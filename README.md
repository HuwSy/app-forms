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

## 🎛️ Inputs

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

## 📝 Text Field Configuration (`text`)

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

## 🔽 Select Field Configuration (`select`)

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

## 📁 File Field Configuration (`file`)

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

# 📤 Output

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

# 🧠 Behaviour Summary

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

# 📄 Example Usage

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
