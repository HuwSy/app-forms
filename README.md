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
