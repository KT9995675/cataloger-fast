# Каталогизатор Google Drive

Виртуальное дерево папок в Google Таблице; физически все файлы — в одной корневой папке Drive (`root_folder_id`). Загрузка через скрытую staging-папку `_Import`.

Локально: `/Users/konstantin/Projects/google-drive-cataloger/`

## Документация

| Файл | Назначение |
|---|---|
| [USER_GUIDE.md](USER_GUIDE.md) | Инструкция для пользователей таблицы |
| [HANDOFF.md](HANDOFF.md) | Handoff для нового чата с AI |
| [SPEC.md](SPEC.md) | Спецификация (частично устарела) |
| [CHANGELOG.md](CHANGELOG.md) | История версий |

## Правила разработки

1. Сначала логика → код по команде **«Пишем»**
2. Одна именованная функция за раз
3. Commit / push — только по команде

## Стек

- Google Sheets — каталог (лист Catalog + скрытый CatalogData)
- Apps Script (`Code.gs`) — логика, HTML-диалоги
- Google Drive API v3 — файлы, права, импорт

## Первичная настройка

1. Google Таблица + Apps Script, `clasp login` → `clasp push`
2. Выполнить **`setupSchema`**
3. Settings → **`root_folder_id`**
4. Развернуть **веб-приложение** (Execute as: Me) → снова `setupSchema` (подтянет URL)
5. Сервисы → **Drive API v3**, переавторизация при новых scopes
6. **Каталогизатор → Первое сканирование** → **Обновить дерево**

Обычным пользователям нужен доступ **Редактор** к папке `_Import` на Drive (назначается автоматически для active Users), **не** к всей `root_folder_id`.

## Статус

**v1.3.0** — staging `_Import`, promote в root, группы `#`, утверждение, корзина, веб-импорт.

**Backlog:** строгая синхронизация прав Catalog ↔ Drive, «Пересчитать всё», UI Users.
