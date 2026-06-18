# app-forms

A simple Angular framework for rapid development originally in AngularJS/JavaScript and progressed to currently run with Angular2+ within an SPFx wrapper giving controls which can use this.spec to determin the field details from the list schema loaded on page load and manipulate this.form for submitting into the list.

Project level installed angular below can be used to avoid additional directories and installs.

```
npm install @angular/cli@22
```
Add to package.json
```
  "scripts": {
    "new": "del package* && ng new --commit=false --routing=false --style=scss --directory .\\"
  },
```
Then
```
npm run new <solution>
```

To generate SSL certs for debugging within localhost and SPFx wrapper
```
npm install -g office-addin-dev-certs
office-addin-dev-certs install --days 3650
```
Then copy these certs into the application directory.

# SharepointChoiceComponent — Inputs & Outputs

The `SharepointChoiceComponent` is a dynamic SharePoint form field renderer supporting text, numbers, dates, users, multi‑choice, file uploads, Outlook/Teams drag‑and‑drop, and more.

---

## Inputs

### **Form & Metadata**

| Input | Type | Description |
|------|------|-------------|
| **`form`** | `SharepointChoiceForm` | Backing form object. Component reads/writes values directly. |
| **`spec`** | `SharepointChoiceList` | Field specification metadata (SharePoint list schema). |
| **`override`** | `string \| SharepointChoiceField` | Overrides field spec. Strings are JSON‑parsed. |
| **`versions`** | `SharepointChoiceForm[]` | Historical versions of the form. |
| **`field`** | `string` | Internal field name (required). |
| **`prefix`** | `string` | Prefix for HTML name attributes for uniqueness. |

---

### **State & Behaviour**

| Input | Type | Description |
|------|------|-------------|
| **`disabled`** | `boolean` | Disables the field UI. |
| **`tooltip`** | `boolean` | Enables/disables tooltip display. |

---

## Text Field Configuration (`text`)

```ts
@Input() text: {
  pattern?: string;
  height?: number;
  width?: number;
  search?: Function;
  select?: Function;
  parent?: any;
}
```

| Property | Description |
|----------|-------------|
| **`pattern`** | Regex validation pattern. |
| **`height`** | Textarea height (px). |
| **`width`** | Minimum width (px). |
| **`search`** | Async search callback for autocomplete. |
| **`select`** | Callback fired when an autocomplete result is selected. |
| **`parent`** | Parent object passed to callbacks. |

---

## Select Field Configuration (`select`)

```ts
@Input() select: {
  none?: string;
  other?: string;
  filter?: Function;
}
```

| Property | Description |
|----------|-------------|
| **`none`** | Text for “none” option. |
| **`other`** | Text for “Other” option. |
| **`filter`** | Function to filter available choices. |

---

## File Field Configuration (`file`)

```ts
@Input() file: {
  extract?: boolean;
  check?: boolean;
  accept?: string;
  download?: boolean;
  uploadonly?: boolean;
  archive?: string;
  view?: number;
  doctypes?: string[];
  doctype?: string;
  notes?: string;
  spec?: SharepointChoiceList;
}
```

| Property | Description |
|----------|-------------|
| **`extract`** | Extract ZIP/EML contents. |
| **`check`** | Show checkboxes for each file. |
| **`accept`** | HTML file input accept filter. |
| **`download`** | Force download instead of preview. |
| **`uploadonly`** | Prevents showing existing files. |
| **`archive`** | Field name used to mark archived files. |
| **`view`** | 0 = all, 1 = not archived, -1 = archived. |
| **`doctypes`** | Allowed document types. |
| **`doctype`** | Field name storing document type. |
| **`notes`** | Field name for notes. |
| **`spec`** | Additional field spec for file metadata. |

---

# Output

### **`change`**

```ts
@Output() change = new EventEmitter<{
  field: string;
  value: any;
  target: HTMLElement;
}>();
```

Emitted whenever the field value changes.

| Property | Description |
|----------|-------------|
| **`field`** | Internal field name. |
| **`value`** | Updated value (primitive or `.results`). |
| **`target`** | DOM element that triggered the change. |

---

# Behaviour Summary

- Automatically initializes correct SharePoint data structures:
  - MultiChoice → `Collection(Edm.String)`
  - UserMulti → `Collection(Edm.Int32)`
  - URL → `{ Url, Description }`
  - Attachments → `{ results: [] }`
- Supports:
  - Autocomplete search (text & user fields)
  - Multi‑select with custom logic
  - Drag‑and‑drop from Outlook, Teams, and local files
  - File extraction (ZIP, EML, MSG)
  - Sorting, filtering, archiving
- Emits refresh events to other component instances via `SharepointChoiceRefresh`.

---

# Example Usage

```html
<app-choice
  [form]="form"
  [spec]="spec"
  [override]="override"
  [field]="'ProjectName'"
  [disabled]="false"
  [text]="{
    search: searchProjects,
    select: onProjectSelected,
    height: 120
  }"
  (change)="onFieldChange($event)"
></app-choice>
```

# SharepointChoiceTable Component Documentation

The `SharepointChoiceTable` component is a standalone Angular table designed for large SharePoint‑style datasets with support for filtering, sorting, paging, column hiding, tabbed views, row selection, and Excel export.

## Component Selector
`<app-table></app-table>`

## Inputs

### allData: SharepointChoiceTabs
Full dataset keyed by tab name.  
Each row is automatically assigned a `_tracking` key for Angular rendering and caching.  
Filtering by `search` is applied inside this setter.

### allCols: SharepointChoiceColumn[]
List of all available columns.  
Changing this clears the internal column cache.

### allTabs: string[]
Optional explicit list of tabs.  
If omitted, tabs are derived from `allData`.

### selectedTab: string | undefined
Currently selected tab.  
Stored in local storage under `Tab`.  
Changing this resets paging to page 1.

### pageSize: number
Number of rows per page.  
Stored in local storage under `Size`.  
Default: `100`.

### loading: boolean
Enables loading state and disables interactions.  
 Clears row cache when changed.

### search: SharepointChoiceRowChild | undefined
Search object applied across all tabs.  
 Triggers:
 - Reload of all data  
 - Reset of caches  
 - Re‑application of stored sort, filter, order, hidden columns  

### prefix: string
Used for local storage keys.  
Defaults to the current page URL.

### tableHeight: string
CSS height of the table container.  
Default: `calc(100vh - 310px)`.

### allEditing: boolean
If true, all cells render in editing mode.

### allowHideColumns: boolean
Enables column hide/show UI.

### showEmptyTabs: boolean
If true, tabs with zero rows are still shown.

### allowSelection: boolean
Enables row selection mode.

### rowClicked?: Function
Callback invoked before all other click events.  
Signature:  
`(row, target) => Promise<void> | void`

### hyperlinkRow?: Function
Optional callback to turn a row into a hyperlink.  
`(row) => string`

### hyperlinkTarget: string
Target attribute for hyperlink rows.  
Default: `_self`.

### export?: SharepointChoiceExportOptions
Declarative export configuration for Excel export.

## Outputs

### selected
Emits `{ data: SharepointChoiceRow[], tab: string | undefined }` when rows are selected.

### cleared
Emits when selection is cleared.

### clicked
Emits `{ row, target }` when a row is clicked (after `rowClicked` if provided).

## Internal State & Behaviour

### Paging
 - `pageNumber` resets page cache  
 - Defaults to page 1

### Sorting (sort: SharepointChoiceSort)
 Stored in local storage under `Sort`.  
 Changing it triggers a debounced refresh.

### Filtering (filter: SharepointChoiceFilter)
 Stored in local storage under `Filter`.  
 Supports:
 - equals  
 - contains  
 - greater  
 - less  

### Column Order (columnOrder: SharepointChoiceOrder)
Stored in local storage under `Order`.  
Changing it clears column cache.

### Hidden Columns (hiddenColumns: SharepointChoiceHide)
 Stored in local storage under `Hide`.  
 Changing it clears column cache.

## Caching Layers

 | Cache | Purpose |
 |-------|---------|
 | _colsCache | Stores computed column lists per tab |
 | _rowsCache | Stores filtered/sorted rows per tab |
 | _pageCache | Stores current page rows |
 | _fieldMapCache | Stores field → values mapping |
 | _nodeCache | Stores rendered node references |

All caches are cleared when relevant inputs change.

## Data Processing Pipeline

1. Tracking keys added to each row (`_tracking`)
2. Search filter applied
3. Tabs filtered (unless `showEmptyTabs`)
4. Caches cleared
5. Change detection triggered

## Lifecycle

### ngOnInit()
Sets up observers and internal state.

### ngOnDestroy()
Cleans up observers and timers.

## Drag & Drop Columns
Internal fields:
 - `_dragColumn`
 - `_dragParent`
 - `_suppressHeaderClick`

Used to support column reordering.

## Excel Export
Uses `Workbook` from `devextreme-exceljs-fork`.  
 Controlled by the `export` input.

## Summary
`SharepointChoiceTable` is a high‑performance Angular table component designed for:
 - Large datasets  
 - Multi‑tab views  
 - Sorting, filtering, paging  
 - Column hiding & reordering  
 - Row selection  
 - Excel export  
 - Custom row click behaviour  
 - Hyperlink rows  

It uses aggressive caching and OnPush change detection to remain fast even with thousands of rows.

# SharePointChoice Utils

## Constructor
- Signature: `constructor(context?: string)`
- Description: Create a SharepointChoiceUtils instance for a given site context (optional). Establishes the PnP SP client and internal context used by other methods.
- Parameters:
  - `context` — optional base URL or site context; if omitted the class attempts to derive it from the current page.

## permissions
- Signature: `permissions(): Promise<SharepointChoicePermission>`
- Description: Return a flat permission object describing the current user's effective permissions.
- Returns: Promise resolving to `SharepointChoicePermission` containing `userId` and a `perms` map (group/security title → boolean).

## hasPermission
- Signature: `hasPermission(object: any, permissions: PermissionKind[]): Promise<boolean>`
- Description: Check whether the current user has any of the provided PermissionKind values on the supplied PnP object (web, list, item).
- Parameters:
  - `object` — PnP object exposing `getCurrentUserEffectivePermissions()`.
  - `permissions` — array of `PermissionKind` to check.
- Returns: Promise resolving to `true` if any permission is present, otherwise `false`.

## search
- Signature: `search(query: string, limit?: number, page?: number, sort?: ISort[], select?: string[], detail?: string[], filter?: string[]): Promise<SearchResults>`
- Description: Execute a SharePoint search query via PnP.
- Parameters:
  - `query` — query text.
  - `limit` — optional row limit (default ~1000).
  - `page` — optional 1-based page number (default 1).
  - `sort` — optional sort descriptors.
  - `select` — optional select properties.
  - `detail` — optional hit/highlighted properties or refiners.
  - `filter` — optional refinement filters.
- Returns: Promise resolving to PnP `SearchResults`.

## fields
- Signature: `fields(listTitle: string): Promise<SharepointChoiceList>`
- Description: Retrieve and normalize list fields metadata for the named list.
- Parameters:
  - `listTitle` — list title.
- Returns: Promise resolving to a `SharepointChoiceList` mapping keyed by field InternalName (includes parsed SchemaXml and Scope property).

## data
- Signature: `data(id: number, listTitle: string): Promise<SharepointChoiceForm>`
- Description: Load a single list item by ID and convert SharePoint values to JS-friendly types for app usage.
- Parameters:
  - `id` — item ID.
  - `listTitle` — list title.
- Returns: Promise resolving to parsed `SharepointChoiceForm`.

## version
- Signature: `version(id: number, listTitle: string, spec?: SharepointChoiceList | null): Promise<SharepointChoiceForm[]>`
- Description: Get version history for a list item and compute changed fields between versions.
- Parameters:
  - `id` — item ID.
  - `listTitle` — list title.
  - `spec` — optional fields metadata to map internal names to display titles.
- Returns: Promise resolving to an array of versions (oldest → newest), each including `ChangedFields`.

## msalApi
- Signature: `msalApi(serverRelativeEndPoint: string, tokenRole: string, httpMethod?: string, jsonPostData?: any, dataType?: string, environment?: string): Promise<any>`
- Description: High-level helper to call mapped backend APIs using MSAL authentication.
- Parameters:
  - `serverRelativeEndPoint` — API endpoint path (server-relative mapping).
  - `tokenRole` — permission scope name used for token acquisition.
  - `httpMethod` — HTTP verb (default `'GET'`).
  - `jsonPostData` — optional JSON body.
  - `dataType` — expected response type (e.g., `'json'`, `'text'`).
  - `environment` — optional environment/release tag.
- Returns: Promise resolving to the API response.

## callApi
- Signature: `callApi(tenancyOnMicrosoft?: string, clientId?: string, permissionScope?: string, apiUrl?: string, httpMethod?: string, jsonPostData?: any, dataType?: string): Promise<any>`
- Description: Generic MSAL-authenticated API caller (can perform token acquisition alone if `apiUrl` omitted).
- Parameters:
  - `tenancyOnMicrosoft` — tenant short name (for authority URL).
  - `clientId` — MSAL client ID.
  - `permissionScope` — scope to request.
  - `apiUrl` — full API URL to call.
  - `httpMethod` — HTTP method (e.g., `'GET'`, `'POST'`).
  - `jsonPostData` — request body for non-GET calls.
  - `dataType` — expected response format (default `'json'`).
- Returns: Promise resolving to parsed response or rejects on error.

## param
- Signature: `param(parameterToReturn: string): string | undefined`
- Description: Read a query-string parameter from the current page location.
- Parameters:
  - `parameterToReturn` — query parameter name.
- Returns: Decoded string value or `undefined` if not present.

## ensurePath
- Signature: `ensurePath(path: string, start: number): Promise<void>`
- Description: Ensure a folder path exists on the site by creating any missing subfolders recursively.
- Parameters:
  - `path` — server-relative or absolute folder path.
  - `start` — start index to control recursion (use 0/2/4 depending on path root).
- Returns: Promise resolving when path exists.

## getRoot
- Signature: `getRoot(list: string): Promise<string>`
- Description: Retrieve the server-relative URL of a list's root folder.
- Parameters:
  - `list` — list title.
- Returns: Promise resolving to the root `ServerRelativeUrl`.

## getFiles
- Signature: `getFiles(path: string, additional?: string): Promise<SharepointChoiceAttachment[]>`
- Description: Get files in a folder and return attachments with metadata, including `ListItemAllFields`.
- Parameters:
  - `path` — server-relative or absolute folder path.
  - `additional` — optional subfolder name appended to `path`.
- Returns: Promise resolving to an array of `SharepointChoiceAttachment`.

## relocateFolder
- Signature: `relocateFolder(source: string, destination: string): Promise<string | null>`
- Description: Move a folder from source to destination server-relative paths.
- Parameters:
  - `source` — source folder server-relative path.
  - `destination` — destination folder server-relative path.
- Returns: Promise resolving to new `ServerRelativeUrl` or `null` if source and destination are equal.

## saveFiles
- Signature: `saveFiles(path: string, additional?: string, url?: { Url: string, Description: string }, files?: { results: SharepointChoiceAttachment[] }, metadata?: SharepointChoiceForm): Promise<void>`
- Description: Save or update files and apply folder/item metadata. Handles folder creation and attachment upload/update.
- Parameters:
  - `path` — target folder path.
  - `additional` — optional subfolder.
  - `url` — optional URL object for metadata.
  - `files` — object with `results` array of attachments to add/update.
  - `metadata` — optional metadata to apply to folder/items.
- Returns: Promise resolving when operation completes.

## save
- Signature: `save(formDataIncIdToUpdate: SharepointChoiceForm, uneditedDataToBuildPatch: SharepointChoiceForm, listTitle: string): Promise<number>`
- Description: Create or update a list item and handle attachments (delete then upload). Computes minimal patch when possible.
- Parameters:
  - `formDataIncIdToUpdate` — data to save; may include `Id` for update.
  - `uneditedDataToBuildPatch` — original data used to compute changes/patch.
  - `listTitle` — list title.
- Returns: Promise resolving to the saved item Id.
