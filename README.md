# app‑forms — Angular Components for SharePoint (SPFx)

**app‑forms** is a lightweight Angular framework for building **dynamic SharePoint list forms**, **high‑performance data tables**, and **PnP‑powered utilities**.  
It works inside **SharePoint Framework (SPFx)** or as a standalone Angular application with some small adaption.

This library includes:

- Dynamic SharePoint form fields (`SharepointChoiceComponent`)
- A high‑performance Angular data table (`SharepointChoiceTable`)
- PnP‑powered SharePoint utilities (`SharepointChoiceUtils`)
- File upload, drag‑and‑drop, metadata extraction, and versioning
- Autocomplete, user pickers, multi‑choice, and complex field types

Designed for **rapid prototyping**, **enterprise apps**, and **large SharePoint datasets**.

---

# Installation & Setup

### Install Angular locally
```bash
npm install @angular/cli@22
```

### Add project scaffolding script
```json
"scripts": {
  "new": "del package* && ng new --commit=false --routing=false --style=scss --directory .\\"
}
```

### Create a new Angular workspace
```bash
npm run new <solution>
```

### Generate SSL certificates for SPFx + localhost debugging
```bash
npm install -g office-addin-dev-certs
office-addin-dev-certs install --days 3650
```

Copy the generated certificates into your Angular project.

---

# SharepointChoiceComponent — Dynamic SharePoint Field Renderer

A fully dynamic Angular component that renders **any SharePoint list field type** based on schema metadata.

Supports:

- Text, number, date, choice, multi‑choice  
- People picker (single/multi)  
- URL fields  
- File uploads (with extraction, metadata, archiving)  
- Drag‑and‑drop from Outlook, Teams, and local files  
- Autocomplete search  
- Version history awareness  

---

# Inputs

## Form & Metadata

| Input | Type | Description |
|-------|------|-------------|
| `form` | `SharepointChoiceForm` | Backing form object. |
| `spec` | `SharepointChoiceList` | SharePoint list schema. |
| `override` | `string \| SharepointChoiceField` | Override field metadata. |
| `versions` | `SharepointChoiceForm[]` | Version history. |
| `field` | `string` | Internal field name. |
| `prefix` | `string` | HTML name prefix. |

## State & Behaviour

| Input | Type | Description |
|-------|------|-------------|
| `disabled` | `boolean` | Disable UI. |
| `tooltip` | `boolean` | Enable/disable tooltips. |

---

# Text Field Configuration

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
| `pattern` | Regex validation. |
| `height` | Textarea height. |
| `width` | Minimum width. |
| `search` | Autocomplete search callback. |
| `select` | Autocomplete selection callback. |
| `parent` | Callback context. |

---

# Select Field Configuration

```ts
@Input() select: {
  none?: string;
  other?: string;
  filter?: Function;
}
```

| Property | Description |
|----------|-------------|
| `none` | “None” label. |
| `other` | “Other” label. |
| `filter` | Filter available choices. |

---

# File Field Configuration

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
| `extract` | Extract ZIP/EML contents. |
| `check` | Show checkboxes for each file. |
| `accept` | HTML file input accept filter. |
| `download` | Force download instead of preview. |
| `uploadonly` | Prevent showing existing files. |
| `archive` | Field name used to mark archived files. |
| `view` | 0 = all, 1 = not archived, -1 = archived. |
| `doctypes` | Allowed document types. |
| `doctype` | Field name storing document type. |
| `notes` | Field name for notes. |
| `spec` | Additional field spec for file metadata. |

---

# Output

### `change` EventEmitter

```ts
@Output() change = new EventEmitter<{
  field: string;
  value: any;
  target: HTMLElement;
}>();
```

| Property | Description |
|----------|-------------|
| `field` | Internal field name. |
| `value` | Updated value (primitive or `.results`). |
| `target` | DOM element that triggered the change. |

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

# SharepointChoiceTable — High‑Performance Angular Table

A fast, enterprise‑grade Angular table for **large SharePoint datasets**, with:

- Multi‑tab views  
- Sorting, filtering, paging  
- Column hiding & reordering  
- Row selection  
- Excel export  
- Hyperlink rows  
- OnPush + multi‑layer caching  

---

# Inputs

| Input | Description |
|-------|-------------|
| `allData` | Full dataset keyed by tab name. |
| `allCols` | List of all available columns. |
| `allTabs` | Optional explicit list of tabs. |
| `selectedTab` | Currently selected tab. |
| `pageSize` | Rows per page. |
| `loading` | Enables loading state. |
| `search` | Search object applied across all tabs. |
| `prefix` | Local storage key prefix. |
| `tableHeight` | CSS height of table container. |
| `allEditing` | Render all cells in edit mode. |
| `allowHideColumns` | Enable hide/show UI. |
| `showEmptyTabs` | Show tabs with zero rows. |
| `allowSelection` | Enable row selection. |
| `rowClicked` | Pre‑click callback. |
| `hyperlinkRow` | Convert row to hyperlink. |
| `hyperlinkTarget` | Link target. |
| `export` | Excel export configuration. |

---

# Outputs

| Output | Description |
|--------|-------------|
| `selected` | Emits selected rows + tab. |
| `cleared` | Emits when selection is cleared. |
| `clicked` | Emits row + target after click. |

---

# Internal Behaviour

- **Paging** resets caches and defaults to page 1  
- **Sorting** stored in local storage (`Sort`)  
- **Filtering** stored in local storage (`Filter`)  
- **Column order** stored in local storage (`Order`)  
- **Hidden columns** stored in local storage (`Hide`)  

### Caching Layers

| Cache | Purpose |
|-------|---------|
| `_colsCache` | Computed column lists per tab |
| `_rowsCache` | Filtered/sorted rows per tab |
| `_pageCache` | Current page rows |
| `_fieldMapCache` | Field → values mapping |
| `_nodeCache` | Rendered node references |

---

# SharepointChoiceUtils — PnP + MSAL Utilities

A utility class wrapping **PnP JS**, **SharePoint REST**, and **MSAL** for:

- Direct PnP access via `utils.sp`
- Access to the current site context via `utils.context`
- Permissions  
- Search  
- List metadata  
- Item loading  
- Version history  
- File operations  
- Folder creation  
- MSAL‑authenticated API calls  

Below is a detailed breakdown.

---

## Constructor
- Signature: `constructor(context?: string)`
- Description: Creates a SharepointChoiceUtils instance for a given site context.  
  The constructor also exposes two useful properties:

### **Exposed Properties**

#### `sp: SPFI`
PnP JS instance pre‑configured with the provided context.  
This allows direct access to the full PnP API surface, for example:

```ts
utils.sp.web.lists.getByTitle('MyList').items.select('Id','Title')();
```

Useful when you need raw PnP operations beyond the high‑level helpers.

#### `context: string`
The site context or base URL passed into the constructor.  
Used internally for:

- Determining folder‑path depth  
- SP request routing  
- Ensuring correct URL resolution  

You may also use it in your own logic when constructing server‑relative paths.

# `permissions()`

### Description  
Returns a flattened permission object describing the current user’s effective SharePoint permissions.

### Inputs  
None.

### Output  
`Promise<SharepointChoicePermission>` containing:  
- `userId` — Current user ID  
- `perms` — Map of permission name → boolean  

---

# `hasPermission(object, permissions)`

### Description  
Checks whether the current user has **any** of the specified `PermissionKind` values.

### Inputs  
- `object` — PnP object (web, list, or item)  
- `permissions` — Array of `PermissionKind` values  

### Output  
`Promise<boolean>`

---

# `search(query, limit?, page?, sort?, select?, detail?, filter?)`

### Description  
Executes a SharePoint search query using PnP.

### Inputs  
- `query` — Search text  
- `limit` — Max rows  
- `page` — Page number  
- `sort` — Sort descriptors  
- `select` — Fields to return  
- `detail` — Highlighted properties / refiners  
- `filter` — Refinement filters  

### Output  
`Promise<SearchResults>`

---

# `fields(listTitle)`

### Description  
Loads and normalises SharePoint list field metadata.

### Inputs  
- `listTitle` — List name  

### Output  
`Promise<SharepointChoiceList>`

---

# `data(id, listTitle)`

### Description  
Loads a single SharePoint list item and converts values into JS‑friendly structures.

### Inputs  
- `id` — Item ID  
- `listTitle` — List name  

### Output  
`Promise<SharepointChoiceForm>`

---

# `version(id, listTitle, spec?)`

### Description  
Loads version history and computes changed fields.

### Inputs  
- `id` — Item ID  
- `listTitle` — List name  
- `spec` — Optional field metadata  

### Output  
`Promise<SharepointChoiceForm[]>`

---

# `msalApi(endpoint, tokenRole, method?, body?, dataType?, environment?)`

### Description  
Calls a backend API using MSAL authentication.

### Inputs  
- `endpoint` — Server‑relative API path  
- `tokenRole` — Permission scope  
- `method` — HTTP verb  
- `body` — JSON body  
- `dataType` — Response type  
- `environment` — Optional environment tag  

### Output  
`Promise<any>`

---

# `callApi(tenant, clientId, scope, url?, method?, body?, dataType?)`

### Description  
Generic MSAL‑authenticated API caller.

### Inputs  
- `tenant` — Tenant short name  
- `clientId` — MSAL client ID  
- `scope` — Permission scope  
- `url` — Full API URL  
- `method` — HTTP verb  
- `body` — JSON body  
- `dataType` — Response format  

### Output  
`Promise<any>`

---

# `param(name)`

### Description  
Reads a query‑string parameter from the current page.

### Inputs  
- `name` — Parameter key  

### Output  
`string | undefined`

---

# `ensurePath(path, start)`

### Description  
Ensures a folder path exists, creating missing subfolders recursively.

### Inputs  
- `path` — Folder path  
- `start` — Index to begin recursion  

### Output  
`Promise<void>`

---

# `getRoot(list)`

### Description  
Gets the server‑relative root folder URL of a list.

### Inputs  
- `list` — List title  

### Output  
`Promise<string>`

---

# `getFiles(path, additional?)`

### Description  
Loads files from a folder and returns attachment metadata.

### Inputs  
- `path` — Folder path  
- `additional` — Optional subfolder  

### Output  
`Promise<SharepointChoiceAttachment[]>`

---

# `relocateFolder(source, destination)`

### Description  
Moves a folder to a new location.

### Inputs  
- `source` — Source folder path  
- `destination` — Destination folder path  

### Output  
`Promise<string | null>`

---

# `saveFiles(path, additional?, url?, files?, metadata?)`

### Description  
Uploads or updates files and applies metadata.

### Inputs  
- `path` — Target folder  
- `additional` — Optional subfolder  
- `url` — URL metadata  
- `files` — `{ results: [...] }` attachments  
- `metadata` — Folder/item metadata  

### Output  
`Promise<void>`

---

# `save(form, original, listTitle)`

### Description  
Creates or updates a SharePoint list item, including attachments.

### Inputs  
- `form` — Data to save  
- `original` — Original data for patching  
- `listTitle` — List name  

### Output  
`Promise<number>` — Saved item ID.

---

# License

MIT
