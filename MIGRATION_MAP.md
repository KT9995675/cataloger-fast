# Migration Map (legacy → Cataloger Fast)

Цель: быстро собирать Fast-реализацию “по одной функции”, переиспользуя проверенные кирпичи из `old/Code.gs`, но не таща legacy-архитектуру (полный render Sheets, префиксы в именах, синхронный Drive).

## 0) Ключевые принципы Fast (коротко)

- Данные: служебные листы `Tree`, `Files`, `CatalogACL`, `DriveGrants` (см. `SPEC.md`).
- Drive по умолчанию закрыт всем; доступ выдаём **на один файл** **одному пользователю** перед открытием.
- Любые операции Drive — фоновые и идемпотентные.

---

## 1) Сопоставление модулей

### 1.1 Drive Batch HTTP (берём как есть)

**Fast модуль:** `drive/batch`

**Legacy доноры (copy-as-is):**
- `executeDriveBatchRequests_`
- `buildDriveBatchHttpBody_`
- `parseDriveBatchResponse_`
- `findDriveBatchResultByContentId_`

**Зачем:** batch copy/metadata/revoke, чанки по 50.

---

### 1.2 Drive: выдача доступа перед открытием (новый модуль, частично из legacy)

**Fast модуль:** `drive/grantOnOpen`

**Legacy доноры (adapt):**
- `listDrivePermissionsWithIds_` (получить `permissionId` для delete)
- `buildDrivePermissionKey_` / `compareDriveRole_` (пригодится для сравнения)
- `resolveDriveFileOpenUrl_` / `normalizeDriveOpenUrl_` (URL для открытия)

**Новый контракт:**
- input: `fileId`, `requesterEmail`
- output: `{ ok, url, permissionId, expiresAt? }`

---

### 1.3 Drive: ночной revoke “выданных за день” (новый модуль)

**Fast модуль:** `drive/revokeDaily`

**Legacy доноры (adapt):**
- `executeDriveBatchRequests_` (batch delete permissions)
- `listDrivePermissionsWithIds_` (если `permissionId` не сохранён — fallback)

**Новый контракт:**
- input: список `(fileId, permissionId)` из `DriveGrants` за сутки
- output: stats `{ revoked, failed }`

---

### 1.4 Импорт copy в фоне (новый модуль; batch copy — из legacy)

**Fast модуль:** `import/copyWorker`

**Legacy доноры:**
- `copyDriveFileDirectToCatalogRootForOwner_` (fallback)
- `buildDriveCopyBatchRequests_` (adapt: сейчас строит по plan-структуре; в Fast будет по job-очереди)

**Новый контракт:**
- очередь job: `(catalog_id, source_file_id, target_folder_id, display_name, status)`
- воркер: copy → записать `file_id` в `Files` → статус `ready`

---

### 1.5 Табличные CRUD для `Tree`/`Files` (новый модуль)

**Fast модуль:** `storage/sheets`

**Legacy доноры (copy-as-is частично):**
- `columnIndexMap_`, `readCell_`, `sanitizeRowValues_`

**Важно:** не использовать `renderCatalogView*`, `buildCatalogTreeState_`, group/formatting.

---

### 1.6 Уникализация имён (можно почти как есть)

**Fast модуль:** `catalog/uniqueNames`

**Legacy доноры (copy-as-is / adapt):**
- `appendAutoSuffixToLabel_` (copy-as-is)
- `allocateUniqueCatalogName_` (adapt: вместо `allNodes` из CatalogData — выборка детей из `Files` по `folder_id`)

---

## 2) Что не переносим

**Do-not-use (в Fast):**
- `renderCatalogViewCore_`, `renderCatalogViewLight_`, `applyCatalogView*`, `createCatalogViewRowGroups_`
- префиксы в имени и всё, что требует Drive-rename при переносе/рекодировании
- синхронный sync ACL “по дереву” на Drive

---

## 3) Предложенный порядок реализации (по одной функции)

Ниже порядок, который даёт “вертикальные срезы” — ранний MVP.

### Итерация 1 — каркас хранения (Sheets)

1. `setupSchemaFast()` — создать листы `Tree`, `Files`, `CatalogACL`, `DriveGrants`
2. `createFolderFast(parentFolderId, label)`
3. `createFileRecordFast(folderId, sourceFileId, displayName)` → статус `pending_copy` или `ready` (в зависимости от режима)

### Итерация 2 — Drive batch кирпичи

4. `driveBatchRequest(requests[])` (обёртка на `executeDriveBatchRequests_`)
5. `driveCopyBatch(jobs[])` — copy 50 файлов, вернуть `newFileId[]`

### Итерация 3 — Импорт copy (быстро + фон)

6. `startImportCopyFast(folderId, sourceFileIds[])`:
   - мгновенно создаёт записи `Files` со статусом `pending_copy`
   - кладёт очередь job в Properties/лист
   - запускает триггер-воркер

7. `processPendingImportJobsFast()`:
   - берёт до 50 pending, делает batch copy
   - пишет `file_id`, ставит `ready`

### Итерация 4 — Grant-on-open (1 файл → 1 requester)

8. `grantFileAccessForOpen(fileId, requesterEmail)`:
   - выдать permission
   - записать `permission_id` в `DriveGrants`
   - вернуть URL

### Итерация 5 — Ночной revoke

9. `revokeDailyGrants()` по `DriveGrants` за сутки

---

## 4) Список “первых доноров” (быстрый доступ)

Из `old/Code.gs`:
- `executeDriveBatchRequests_`, `parseDriveBatchResponse_`, `buildDriveBatchHttpBody_`
- `listDrivePermissionsWithIds_`
- `resolveDriveFileOpenUrl_`
- `appendAutoSuffixToLabel_`
- `readCell_`, `columnIndexMap_`

