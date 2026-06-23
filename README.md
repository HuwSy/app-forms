# app-forms

Angular components and utilities for SharePoint forms, tables, and file-driven workflows.

This repository contains three related pieces:

- `sharepoint-choice`: the Angular library
- `sample`: a standalone Angular sample app that consumes the library
- `angular-wrapper`: an SPFx web part that loads a bundled Angular app from a document library folder and renders it on the page
- `AngularJS`: the start of this journey and kept only for nostalgia

The library is aimed at SharePoint-hosted Angular experiences, including:

- dynamic field rendering from SharePoint list metadata
- editable, cached data tables for large datasets
- PnP-powered helpers for list data, files, folders, search, and MSAL-backed API calls
- attachment handling with extraction, metadata, and archive workflows

## Packages In This Repo

### `sharepoint-choice`

Exports:

- `SharepointChoiceComponent`
- `SharepointChoiceTable`
- `SharepointChoiceUtils`
- `SharepointChoiceLogging`
- `SharepointChoiceRefresh`
- shared models and interfaces

Public API entry point:

```text
sharepoint-choice/src/public-api.ts
```

### `sample`

The sample app demonstrates how to:

- bootstrap a standalone Angular component into a page
- render `SharepointChoiceComponent`
- render `SharepointChoiceTable`
- use `SharepointChoiceUtils` against a SharePoint site

The sample component selector is:

```html
<app-sample></app-sample>
```

### `angular-wrapper`

This is an SPFx web part named `Angular Wrapper`. It loads client-side assets from a folder and injects a chosen selector into the page. Its key web part properties are:

- `Folder`: URL to the script folder
- `Selector`: Angular selector tag name
- `Additional`: additional attributes to place on the element
- `ESBuild used`: whether the target bundle is emitted in the newer Angular build format

## Repository Setup

### Prerequisites

- Node.js compatible with the package being built
- npm
- Angular CLI 22 for the Angular library and sample app

Install Angular CLI locally where needed:

```bash
npm install @angular/cli@22
```

### Build the library

From `sharepoint-choice`:

```bash
npm install
npm run build
```

This packages the library to:

```text
sharepoint-choice/dist/sharepoint-choice
```

### Build the sample app

From `sample`:

```bash
npm install
npm run spc
npm run bundle
```

Notes:

- `npm run spc` builds the local `sharepoint-choice` package first
- the sample app uses Angular's `@angular/build:application` builder
- with that builder, browser assets are emitted under the `browser` output directory by default

Given the current `sample/angular.json`, the bundled browser assets are expected under:

```text
sample/dist/sample/browser
```

### Local HTTPS for SharePoint-hosted debugging

If you want to test against SharePoint pages over localhost, generate a dev certificate:

```bash
npm install -g office-addin-dev-certs
office-addin-dev-certs install --days 3650
```

Then place the generated certificate files where the sample app expects them:

- `sample/localhost.crt`
- `sample/localhost.key`

## Running Through SPFx With `angular-wrapper`

If you want to host the built Angular app in SharePoint and execute it on a page through SPFx, this repo already contains the wrapper pattern.

### What the wrapper loads

For `ESBuild used = true`, the wrapper loads these files from the configured folder:

- `styles.css`
- `polyfills.js`
- `main.js`

For `ESBuild used = false`, it instead expects:

- `polyfills.js`
- `runtime.js`
- `main.js`
- `styles.css`

The sample app uses Angular's application builder, so the correct wrapper setting for the sample is:

- `ESBuild used = true`

### Recommended deployment flow

1. Build the library in `sharepoint-choice`.
2. Build the sample app in `sample`.
3. Upload the emitted browser assets to a SharePoint document library folder.
4. Add the `Angular Wrapper` web part to a page.
5. Point the web part at the uploaded folder.
6. Set the selector to the Angular root component you want to bootstrap.

### Correct folder naming

Because the sample build uses the default Angular browser output subfolder, you have two valid deployment options.

#### Option A: Upload the `browser` folder contents into your target SharePoint folder

Example upload target:

```text
SiteAssets/sample/v1
```

In that case, set the wrapper property values to:

- `Folder`: `SiteAssets/sample/v1`
- `Selector`: `app-sample`
- `ESBuild used`: checked

#### Option B: Upload the whole `dist/sample` structure and preserve the `browser` folder

Example upload target:

```text
SiteAssets/sample/v1/browser
```

In that case, set the wrapper property values to:

- `Folder`: `SiteAssets/sample/v1/browser`
- `Selector`: `app-sample`
- `ESBuild used`: checked

### Tenant app catalog note

If you upload the bundle into the tenant app catalog site's `SiteAssets` library, the same pattern applies. Use the folder that directly contains `main.js`, `polyfills.js`, and `styles.css`.

Examples:

- site-relative from the app catalog site: `SiteAssets/sample/v1/browser`
- server-relative: `/sites/AppCatalog/SiteAssets/sample/v1/browser`
- absolute: `https://tenant.sharepoint.com/sites/AppCatalog/SiteAssets/sample/v1/browser`

The wrapper accepts all three forms and normalizes them before loading the assets.

### Selector and attributes

For the sample app:

- `Selector`: `app-sample`

You can also pass extra attributes through the wrapper's `Additional` property. The wrapper already injects a `context` attribute containing the current SharePoint web URL.

## `SharepointChoiceComponent`

`SharepointChoiceComponent` is a standalone Angular component that renders SharePoint field UI from field metadata and a backing form object.

### Supported rendered field types

The component currently contains renderers for:

- `Boolean`
- `MultiChoice`
- `Choice`
- `Integer`
- `Number`
- `Currency`
- `DateTime`
- `Text`
- `Lookup`
- `Note`
- `Geolocation`
- `URL`
- `Attachments`
- `User`
- `UserMulti`

Notes:

- `Lookup` includes built-in SharePoint list search against `LookupList` and `LookupField`, and keeps a display label visible for loaded or disabled values
- `Currency` uses locale-aware display and parsing based on the field LCID, with shared locale and currency-region maps exported from `sharepoint-choice.models.ts`
- `Geolocation` renders latitude and longitude inputs and shows a map link when both coordinates are present
- the shared field model still includes `LookupMulti`, but this component does not currently contain a dedicated render case for it
- rich text note fields are supported through `@bobbyquantum/ngx-editor` when the field metadata indicates rich text

### Locale and country coverage for currency

Currency formatting is not hardcoded to a single symbol. The component maps SharePoint LCIDs to BCP 47 locales and then maps locale regions to currency codes.

The current built-in coverage includes a broad set of common SharePoint locales and countries, including examples such as:

- United States and United Kingdom
- most common Euro-area locales such as Germany, France, Spain, Italy, Portugal, Belgium, Slovakia, Slovenia, Estonia, Latvia, and Lithuania
- Switzerland, Norway, Sweden, Denmark, Poland, Czech Republic, Hungary, Romania, Croatia, Serbia, Russia, and Ukraine
- Saudi Arabia, Israel, Turkey, Pakistan, India, Thailand, Malaysia, Indonesia, Kazakhstan, Vietnam, China, Taiwan, Hong Kong, Singapore, and Macau
- Brazil, Japan, and South Korea

If an LCID or region is not in the built-in map, the component falls back to `en-US` and `USD`.

### Inputs

| Input | Type | Description |
|-------|------|-------------|
| `form` | `SharepointChoiceForm` | Backing object for the field values. |
| `spec` | `SharepointChoiceList | undefined` | List field metadata keyed by internal name. |
| `override` | `string | SharepointChoiceField | undefined` | Metadata override for the current field. String values are parsed as JSON. |
| `disabled` | `boolean` | Disables the control. |
| `versions` | `SharepointChoiceForm[] | undefined` | Version history for display helpers. |
| `prefix` | `string` | Prefix added to generated form control names. |
| `field` | `string` | Internal field name to bind. |
| `text` | `object` | Text and note field configuration. |
| `select` | `object` | Choice field configuration. |
| `file` | `object` | Attachment and document field configuration. |

`tooltip` is internal component state, not a public input.

### `text` configuration

```ts
{
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
| `pattern` | Regex validation pattern for text input. |
| `height` | Textarea height in pixels. |
| `width` | Minimum width in pixels. |
| `search` | Async callback for autocomplete results. |
| `select` | Callback when an autocomplete item is chosen. |
| `parent` | Optional callback context object. |

### `select` configuration

```ts
{
  none?: string;
  other?: string;
  filter?: Function;
}
```

| Property | Description |
|----------|-------------|
| `none` | Label for a null option. |
| `other` | Label for an "Other" option. |
| `filter` | Function used to filter available choices. |

### `file` configuration

```ts
{
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
| `extract` | Extract ZIP, EML, or MSG contents where supported. |
| `check` | Show checkboxes against files. |
| `accept` | File input accept filter. |
| `download` | Force download links instead of opening a new tab. |
| `uploadonly` | Hide existing files and only allow upload interactions. |
| `archive` | Field name used to mark archived files. |
| `view` | `0` all, `1` not archived, `-1` archived. |
| `doctypes` | Allowed document type values. |
| `doctype` | Field name that stores document type. |
| `notes` | Field name used for file notes. |
| `spec` | Additional field metadata used for file-level editing. |

### Output

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
| `value` | Updated value. For multi-value structures, `.results` is emitted where applicable. |
| `target` | The component host element, not necessarily the inner native control that triggered the update. |

### Behavior summary

- initializes SharePoint-shaped values for fields such as `MultiChoice`, `UserMulti`, `URL`, and `Attachments`
- supports autocomplete on text and people fields
- supports drag and drop from local files, Teams library links, and Outlook add-in flows
- supports file extraction for ZIP, EML, and MSG uploads
- emits refresh notifications through `SharepointChoiceRefresh` so sibling instances using the same backing references can update under `OnPush` change detection

## `SharepointChoiceTable`

`SharepointChoiceTable` is a standalone Angular table component for SharePoint-style datasets.

### Features

- multiple tabs over one table surface
- persisted sort, filter, hidden-column, tab, and page-size state via local storage
- paging
- row selection with emitted selected rows
- inline editing using `SharepointChoiceComponent`
- column hide/show, resize, and reorder
- hyperlink rows
- custom cell renderers
- Excel export
- memoized row, column, page, and renderer caching

### Inputs

| Input | Type | Description |
|-------|------|-------------|
| `allData` | `SharepointChoiceTabs` | Full dataset keyed by tab name. |
| `allCols` | `SharepointChoiceColumn[]` | Column definitions. |
| `allTabs` | `string[]` | Optional explicit tab order. |
| `selectedTab` | `string | undefined` | Active tab. |
| `pageSize` | `number` | Rows per page. |
| `loading` | `boolean` | Loading state. |
| `search` | `SharepointChoiceRowChild | undefined` | External search filter applied before per-column filtering. |
| `prefix` | `string` | Prefix for local storage keys. |
| `tableHeight` | `string` | Table container height. |
| `allEditing` | `boolean` | Render editable cells in edit mode by default. |
| `allowHideColumns` | `boolean` | Enable hide/show UX. |
| `showEmptyTabs` | `boolean` | Include tabs with no rows. |
| `allowSelection` | `boolean` | Enable row selection UI. |
| `rowClicked` | `Function | undefined` | Optional row click callback. |
| `hyperlinkRow` | `Function | undefined` | Optional row-to-URL callback. |
| `hyperlinkTarget` | `string` | Link target for hyperlink rows. |
| `export` | `SharepointChoiceExportOptions | undefined` | Export configuration. |

### Outputs

| Output | Type | Description |
|--------|------|-------------|
| `selected` | `EventEmitter<{ data, tab }>` | Emits selected rows for the active tab. |
| `cleared` | `EventEmitter<void>` | Emits when table state is cleared. |
| `clicked` | `EventEmitter<{ row, target }>` | Emits row click data when no callback overrides it. |

### Internal caching

| Cache | Purpose |
|-------|---------|
| `_colsCache` | Computed visible columns per tab |
| `_rowsCache` | Filtered and sorted rows per tab |
| `_pageCache` | Current page slice |
| `_fieldMapCache` | Parsed dot-path field segments |
| `_nodeCache` | Cached renderer node output |

## `SharepointChoiceUtils`

`SharepointChoiceUtils` wraps PnP JS, SharePoint REST behavior, and MSAL-backed API access.

### Constructor

```ts
constructor(context?: string)
```

If no context is supplied, the utility attempts to derive one from the current SharePoint page.

### Exposed properties

| Property | Type | Description |
|----------|------|-------------|
| `context` | `string` | Resolved site URL used for SharePoint operations. |
| `sp` | `SPFI` | PnP JS instance configured against `context`. |

### Method summary

| Method | Returns | Purpose |
|--------|---------|---------|
| `permissions()` | `Promise<SharepointChoicePermission>` | Builds a flattened permission or group map for the current user. |
| `hasPermission(object, permissions)` | `Promise<boolean>` | Checks effective SharePoint permissions against `PermissionKind` values. |
| `search(query, limit?, page?, sort?, select?, detail?, filter?)` | `Promise<SearchResults>` | Runs SharePoint search. |
| `fields(listTitle)` | `Promise<SharepointChoiceList>` | Loads and normalizes list field metadata. |
| `data(id, listTitle)` | `Promise<SharepointChoiceForm>` | Loads one list item and normalizes values. |
| `version(id, listTitle, spec?)` | `Promise<SharepointChoiceForm[]>` | Loads version history and computes changed fields. |
| `msalApi(endpoint, tokenRole, method?, body?, dataType?, environment?)` | `Promise<any>` | Calls a mapped API using MSAL. |
| `callApi(tenant, clientId, scope, url?, method?, body?, dataType?)` | `Promise<any>` | Generic MSAL-authenticated API caller. |
| `save(form, original, listTitle)` | `Promise<number>` | Creates or updates a SharePoint list item and attachments. |
| `param(name)` | `string | undefined` | Reads a query-string parameter. |
| `ensurePath(path, start)` | `Promise<void>` | Ensures a folder path exists. |
| `getRoot(list)` | `Promise<string>` | Returns a list root folder URL. |
| `getFiles(path, additional?)` | `Promise<SharepointChoiceAttachment[]>` | Loads files and list item metadata from a folder. |
| `relocateFolder(source, destination)` | `Promise<string | null>` | Moves a folder. |
| `saveFiles(path, additional?, url?, files?, metadata?)` | `Promise<void>` | Saves uploaded files and applies metadata. |

### Important note on `permissions()`

`permissions()` does not compute full effective SharePoint permissions in the same way as `hasPermission()`. It builds a flat `perms` object from:

- current SharePoint group memberships
- optional entries from a hidden `Security` list when present, this can be used with a list item per permission and secured to only that permission to facilitate nested permission discovery which can happen in some tenancies.

Use `hasPermission()` when you need an object-level effective-permission check against SharePoint permission kinds.

## License

MIT
