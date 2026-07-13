# Handoff — перезапуск чата / продолжение работы

> Компактная выжимка для нового диалога с AI. Архитектура — [SPEC.md](SPEC.md). История — [CHANGELOG.md](CHANGELOG.md). Для пользователей — [USER_GUIDE.md](USER_GUIDE.md).

## Проект

**Каталогизатор Google Drive** — виртуальное дерево папок в Google Таблице; физически все файлы лежат в **одной корневой папке** Drive (`root_folder_id`).

- Локально: `/Users/konstantin/Projects/google-drive-cataloger/`
- Деплой: `clasp push` из корня
- Script ID — в `.clasp.json` (из Apps Script, **не** ID таблицы)
- Включить: [Apps Script API](https://script.google.com/home/usersettings) + **Сервисы → Drive API v3**
- **Версия в коде:** `SCHEMA_VERSION = '1.4.2'` (`Code.gs`)
- **Последний commit в git:** `8fe0419` — v1.3.1 (integrity check)
- **Локально:** ~2600 строк diff, **не закоммичено**; последний `clasp push` — сессия 13.07.2026 (~21:24)

## Правила работы (от владельца) — обязательно

1. **Сначала обсуждаем** предложения и архитектуру.
2. Код **только по команде «Пишем»**.
3. Одна именованная функция за раз, с проверкой (по возможности).
4. **Когда договорённость зафиксирована** — сразу обновлять [SPEC.md](SPEC.md) (раздел «Актуализация v1.4»); при необходимости — этот файл. Не ждать commit.
5. `CHANGELOG.md` — при релизе / по команде на commit.
6. Git commit / push — **только по команде**.
7. Cursor rule: `.cursor/rules/spec-on-agreement.mdc`

## OAuth scopes (`appsscript.json`)

- `drive`, `spreadsheets.currentonly`, `userinfo.email`, `script.container.ui`
- `script.external_request` — веб-импорт (UrlFetchApp), для **редакторов**
- `script.scriptapp` — фоновые триггеры (утверждение, импорт ACL)

После добавления scope — **переавторизация** скрипта.

---

## Что сделано в сессии 13.07.2026 (v1.4.0 локально)

### CatalogIndex + Catalog view

- Скрытый лист **CatalogIndex** (строка N view = строка N index → `data_row_index` в CatalogData).
- Catalog: **11 колонок** (без скрытого `_id`), колонка **Размер** (MB), полосы строк.
- Выбор строки: номер Catalog → CatalogIndex → одна строка CatalogData.

### Ссылки на файлы (исправлено)

- **Баг:** RichText-ссылки сбрасывались `applyCatalogViewFolderBold_` / полосами, если ставились до них.
- **Фикс:** в `renderCatalogViewCore_` порядок — bold/полосы **до** `applyCatalogViewTreeNameCells_` (**ссылки последними**).
- `renderCatalogViewLight_` — без `skipDriveLinks` (ссылки после корзины сохраняются).

### Корзина

- Виртуальная, `rootIds` из меню (не читать выделение после `Ui.alert`).
- `renderCatalogViewLight_` после отправки/восстановления.

### Импорт — архитектура v1.4.1 (владелец каталога)

| Было | Стало |
|---|---|
| N × `google.script.run` → `importCatalogDriveEntry` | **`importCatalogDriveBatch`** |
| «Подготовка» + «Импорт» = 2 RPC для владельца | **1 RPC** (подготовка пропущена) |
| `buildCatalogImportContext_` → sync `_Import` | **`buildCatalogImportContextLight_`** (без `_Import`) |
| copy → `_Import` → promote | **copy сразу в `root_folder_id`** |
| `syncDrivePermissionsFromJson_` на файл | **`processPendingImportDriveSync_`** (~2 с, триггер) |
| `appendRow` × N | **`appendCatalogDataRowsBatch_`** |
| Полный `renderCatalogViewCore_` | **инкрементальная вставка** до 10 файлов, иначе `renderCatalogViewLight_` |
| 1 файл: batch HTTP | **один `Drive.Files.copy`** |
| 2–50 файлов | **Drive Batch API** (`executeDriveBatchRequests_`, до 50 за HTTP) |
| 51–100 файлов | **чанки по 50** — несколько RPC, прогресс между чанками |

**Ключевые функции:** `importCatalogDriveBatch`, `importCatalogDriveBatchCore_`, `prepareCopyImportPlanItem_`, `executeDriveCopyPlansInBatches_`, `copyDriveFileDirectToCatalogRootForOwner_`, `debugImportSingleDriveFile` (есть в коде, диалог использует batch).

**Диалог (`ImportDialog.html`):**
- Владелец: без `prepareDriveImportExecution`; 1 файл → один batch-вызов + **тайминги** в textarea; **не автозакрывается**.
- 2+ файлов: чанки по 50, progress bar обновляется **между** RPC (polling через CacheService **не работает** — GAS сериализует вызовы из диалога).
- Редакторы: по-прежнему `prepareDriveImportExecution` + `importCatalogDriveEntry` по файлу + webapp + `_Import`.

**Константы:** `DRIVE_BATCH_MAX = 50`, `IMPORT_DRIVE_BATCH_CHUNK_SIZE` (legacy, в контексте — `DRIVE_BATCH_MAX`).

### Утверждение

- Фоновый триггер `processPendingApproveDriveSync_` (~2 с).

### Прочее

- Меню перегруппировано; служебные пункты в подменю.
- Drive API: `addParents`/`removeParents` только в 4-м аргументе `Drive.Files.update`.

---

## v1.4.2 — трёхфазный импорт (локально)

### Импорт copy (владелец) — `startCatalogDriveImport`

1. **Фаза 1 (~10–20 с):** Drive Batch scan имён/mime + read CatalogData (родитель, соседи)
2. **Фаза 2 (~3–6 с):** append CatalogData (`file_id` пустой), view incremental ≤10 / defer
3. **Фаза 3 (фон ~45–120 с):** триггер `processPendingDriveImportJob_`, batch copy, ACL, lock

- Диалог: один RPC `startCatalogDriveImport`, отчёт с таймингами, без автозакрытия
- Блокировка: `import_in_progress` + `assertCatalogImportNotBusy_` в мутациях
- Сброс: **Служебные → Сбросить зависший импорт**
- Редакторы: без изменений (`importCatalogDriveEntry` + webapp)
- `SCHEMA_VERSION = 1.4.2`

### Проверить

1. Импорт 1 файла — тайминги в отчёте, фон ~2–5 с
2. Импорт 100 — каталог сразу, Drive в фоне; блокировка меню
3. CHANGELOG / commit — по команде

## Типичные ошибки

| Симптом | Причина / решение |
|---|---|
| Нет ссылок на файлы | Порядок отрисовки; «Обновить дерево» |
| Импорт «висит» без прогресса | Внутри одного RPC (до 50 copy) UI не обновляется — норма |
| Progress polling не работал | `google.script.run` из диалога **один за раз** — исправлено: прогресс между чанками |
| Медленная «подготовка» у владельца | Был лишний RPC + `_Import` sync — убрано для владельца |
| Корзина «не работает» | `rootIds` из меню, не выделение после alert |
| Webapp 404 | redeploy + `setupSchema`; владелец обходит HTTP |

## Backlog (без изменений)

- strict reconcile ACL, миграция root ACL, recovery `_Import`, UI Users
- Фаза B импорта: batch для редакторов через webapp

---

## Файлы изменены (git diff)

```
Code.gs           — ~2500 строк (+CatalogIndex, batch import, links fix, …)
ImportDialog.html — batch UI, progress, result textarea
SPEC.md           — Актуализация v1.4 + v1.4.1 импорт
HANDOFF.md        — этот файл
USER_GUIDE.md     — мелкие правки
.cursor/rules/spec-on-agreement.mdc — правило обновления SPEC
```

---

## Промпт для нового чата (скопировать)

```
Проект: google-drive-cataloger
Путь: /Users/konstantin/Projects/google-drive-cataloger/
Прочитай HANDOFF.md, SPEC.md (раздел «Актуализация v1.4»), CHANGELOG.md.

Правила: сначала обсуждаем; код только по «Пишем»; договорённость → обновлять SPEC.
Git commit — только по команде.

Состояние:
- SCHEMA_VERSION 1.4.0 в Code.gs, clasp push сделан, git на 8fe0419 (v1.3.1), ~2600 строк не закоммичено.
- Сделано: CatalogIndex, fix ссылок (links после bold), batch-импорт для владельца
  (importCatalogDriveBatch, Drive Batch API, light context, 1 RPC, чанки по 50).
- Диалог: отчёт с таймингами, без автозакрытия.

Проверить / продолжить:
1. Тайминги импорта 1 файла (отчёт в диалоге) — цель <5 с
2. Импорт 100 файлов — progress 50/100
3. CHANGELOG v1.4, commit по команде
4. [твоя задача]
```

---

## Контекст сессии (13.07.2026)

1. clasp push, тест ссылок после корзины, импорт copy — commit отложен.
2. **Ссылки пропали** — fix порядка отрисовки (Пишем).
3. **Импорт медленный** (>10 с/файл) — обсуждение batch, Drive Batch API.
4. **100 файлов зависли** — один RPC >6 мин; добавлены чанки по 10→50.
5. **Progress не работал** — polling блокируется GAS; прогресс между чанками.
6. **1 файл всё ещё медленно** — убраны prepare + `_Import` sync для владельца; отчёт без автозакрытия.
7. Пользователь: импорт от **владельца**, режим **копирование** чужих файлов.

**Следующий шаг:** получить тайминги 1 файла из textarea отчёта → при необходимости ускорить `load_catalog` (большой CatalogData) или `view_*` (fallback на light render).
