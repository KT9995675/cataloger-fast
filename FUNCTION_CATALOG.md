# Каталог функций (legacy → донор для Fast)

Этот документ — навигатор по функциям из `old/Code.gs` (репозиторий `google-drive-cataloger`).  
Цель: быстро находить готовые кирпичи и понимать, что можно переиспользовать в `Cataloger Fast`.

## Легенда

- **copy-as-is**: можно брать почти без правок
- **adapt**: брать как основу, но переписать под новую модель (`Tree`/`Files`, grant-on-open)
- **do-not-use**: логика противоречит Fast-модели или тянет лишние зависимости (полный render/Drive sync в критическом пути)

## 1) Утилиты и базовые примитивы (copy-as-is)

- `readCell_`: нормализация значения ячейки → string/''.
- `sanitizeFolderLabel_`: безопасное имя (обрезка/нормализация) для UI/каталога.
- `generateId_`: UUID (можно заменить на Utilities.getUuid напрямую).
- `columnIndexMap_`: map заголовок → номер колонки.
- `arraysEqualAsSets_`: сравнение массивов как множеств.

## 2) Drive Batch HTTP (copy-as-is)

Использовать как основу для любых операций Drive в фоне (чанки по 50).

- `executeDriveBatchRequests_`: отправка multipart batch к `https://www.googleapis.com/batch/drive/v3`.
- `buildDriveBatchHttpBody_`: сборка тела batch-запроса.
- `parseDriveBatchResponse_`: парсинг multipart ответа.
- `findDriveBatchResultByContentId_`: поиск результата sub-request.
- `copyDriveFileDirectToCatalogRootForOwner_`: одиночный `Drive.Files.copy` (fallback).

## 3) Drive: ссылки/открытие (adapt)

В Fast-модели ссылка может требовать grant-on-open.

- `resolveDriveFileOpenUrl_`: получение/нормализация `webViewLink` → пригодится как URL-фабрика.
- `normalizeDriveOpenUrl_`, `getDriveFileOpenUrlFallback_`: нормализация `/view`→`/edit`.

## 4) Drive: ACL и синхронизация (adapt / do-not-use частично)

Fast-модель: Drive “закрыт всем”, доступ выдаётся **на файл** и **только requester’у** перед открытием.

Полезные кирпичи:

- `listDrivePermissionsWithIds_`: получить permissions с `permissionId` (важно для быстрого revoke).
- `buildDrivePermissionKey_`, `compareDriveRole_`: сравнение/ключи.

Логику “sync дерева прав на Drive” из legacy **не переносить напрямую**:

- `syncDrivePermissionsFromJson_`, `applyPermissionsFromJson_`, `syncDrivePermissionsForApproval_` — **adapt** (только для закрытия базовых прав и точечных grant/revoke).

## 5) Импорт/копирование (adapt)

В Fast: “создать записи быстро” + фоновые Drive-copy/ACL.

Полезные части:

- `buildDriveCopyBatchRequests_`: batch `files.copy` на 50.
- `executeDriveCopyPlansInBatches_`: стратегия batch + fallback на seq.
- `prepareCopyImportPlanItem_`: подготовка unique name + target (в Fast заменится на `Tree`/`Files` модель).
- `scanDriveImportFileMetadata_` + `buildDriveMetadataBatchGetRequests_`: batch scan `name/mimeType`.

Legacy UI/планирование по CatalogData — **do-not-use** (в Fast будет другая модель хранения).

## 6) Табличный движок дерева и view (do-not-use для Fast)

В legacy это основная масса времени (render, группы строк, RichText ссылки). В Fast UI будет webapp/иной слой.

- `renderCatalogViewCore_`, `renderCatalogViewLight_`, `buildCatalogTreeState_`, `createCatalogViewRowGroups_`,
  `applyCatalogView*`, `insertCatalogViewNodeAfterCreate_`, `CatalogIndex`-связка — **do-not-use** как целевой дизайн.

Можно использовать только как “референс”, если понадобится временный Sheets-view.

## 7) Корзина/перенос/копирование (adapt)

Полезно как бизнес-правила (конфликты имён, уникализация), но реализацию менять под новую схему:

- `allocateUniqueCatalogName_`, `appendAutoSuffixToLabel_`: **copy-as-is** (механика уникализации).
- `isSiblingCatalogNameTaken_`: **copy-as-is** (проверка коллизий на уровне parent).

Остальное (перенос/копирование поддеревьев, рекодирование, префиксы в имени) — **do-not-use**.

## 8) Группы (adapt)

Fast: группы — только UI-удобство выбора пользователей, на Drive не отражаются.

- `readCatalogGroupsMap_`, `readCatalogGroupMembers_`, `normalizeCatalogGroupMemberEmails_`: **adapt**
- `resyncDriveForCatalogGroup_`: **do-not-use** (Drive не синхронизируем по группам)

## 9) Integrity check / repair (adapt)

Если в Fast останется “проверка целостности” — можно вытащить подход и структуру отчёта.

- `runCatalogIntegrityCheckCore_`, `saveIntegrityReportSnapshot_`, `loadIntegrityReportSnapshot_`: **adapt**

---

## Быстрый “чеклист” переиспользования

В Fast почти наверняка пригодится:

- Drive Batch: `executeDriveBatchRequests_` (+ парсинг)
- ACL list/revoke: `listDrivePermissionsWithIds_`
- уникализация: `allocateUniqueCatalogName_`
- нормализация/утилиты: `readCell_`, `sanitizeFolderLabel_`, `generateId_`

