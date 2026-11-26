import {
  Component,
  OnInit,
  OnDestroy,
  Input,
  ElementRef,
  ChangeDetectorRef,
  ErrorHandler,
  Output,
  EventEmitter,
  ChangeDetectionStrategy,
} from "@angular/core";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import { Editor, NgxEditorModule, Toolbar } from "@bobbyquantum/ngx-editor";
import { Plugin } from "prosemirror-state";
import MsgReader from "@kenjiuno/msgreader";
import PostalMime, { type Attachment as PostalMimeAttachment } from "postal-mime";
import { loadAsync } from "@turbowarp/jszip";
import { Subject, Subscription } from "rxjs";
import { debounceTime } from "rxjs/operators";
import { SharepointChoiceUtils, MapErrors } from "./sharepoint-choice.utils";
import { CommonModule } from "@angular/common";
import { FormsModule } from "@angular/forms";
import { SharepointChoiceLogging } from "./sharepoint-choice.logging";
import {
  SharepointChoiceTextOptions,
  SharepointChoiceSelectOptions,
  SharepointChoiceFileOptions,
  SharepointChoiceForm,
  SharepointChoiceList,
  SharepointChoiceField,
  SharepointChoiceSearchResult,
  SharepointChoiceChangeEvent,
  SharepointChoiceAttachment,
  LCID_TO_BCP47,
  CURRENCY_BY_REGION,
  MAX_PASTED_IMAGE_SIZE_BYTES,
  MAX_PASTED_IMAGE_SIZE_MB,
} from "./sharepoint-choice.models";
import { SharepointChoiceRefresh } from "./sharepoint-choice.refresh";

interface SharepointChoiceUser {
  Id: number;
  Title: string;
  LoginName: string;
}

interface SharepointChoiceLookupLabel {
  id: number;
  label: string;
}

interface SharepointChoiceWindow extends Window {
  Office?: typeof Office;
  chrome?: unknown;
}

interface OutlookAttachmentResponse {
  lastModifiedDateTime: string;
  contentBytes: string;
}

interface SharePointFolderProperties {
  // SharePoint folder property names vary by library/version; callers validate the property they consume.
  [propertyName: string]: unknown;
}

interface SharePointFileListItem {
  File: {
    TimeCreated: string;
    TimeLastModified: string;
    ServerRelativeUrl: string;
    Name: string;
  };
}

type SharePointFileItemKey = [unknown, string, string, string | number];

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function readResults(value: unknown): unknown[] | undefined {
  const record = isRecord(value) ? value : undefined;
  return Array.isArray(record?.["results"]) ? record["results"] : undefined;
}

function isSharepointChoiceSearchResult(value: unknown): value is SharepointChoiceSearchResult {
  return (
    isRecord(value) &&
    (value["Id"] === undefined || (typeof value["Id"] === "number" && Number.isSafeInteger(value["Id"])))
  );
}

function isSharepointChoiceSearchResults(value: unknown): value is SharepointChoiceSearchResult[] {
  return Array.isArray(value) && value.every(isSharepointChoiceSearchResult);
}

function setResults(form: SharepointChoiceForm, field: string, results: unknown[]): void {
  const record = isRecord(form[field]) ? form[field] : {};
  record["results"] = results;
  form[field] = record;
}

function isSharepointChoiceAttachment(value: unknown): value is SharepointChoiceAttachment {
  return isRecord(value) && typeof value["FileName"] === "string";
}

function isSharepointChoiceUser(value: unknown): value is SharepointChoiceUser {
  return (
    isRecord(value) &&
    typeof value["Id"] === "number" &&
    Number.isSafeInteger(value["Id"]) &&
    typeof value["Title"] === "string" &&
    typeof value["LoginName"] === "string"
  );
}

function normalizePeoplePickerResults(value: unknown): SharepointChoiceUser[] {
  if (!Array.isArray(value)) throw new TypeError("The SharePoint people search response must be an array.");

  const users: SharepointChoiceUser[] = [];
  for (const person of value) {
    if (!isRecord(person)) throw new TypeError("The SharePoint people search response contains an invalid entry.");
    const entityData = isRecord(person["EntityData"]) ? person["EntityData"] : undefined;
    if (!entityData) continue;

    const email = entityData["Email"];
    const loginName = person["Key"];
    const title = person["DisplayText"];
    const rawId = entityData["SPUserID"];
    if (email === undefined || email === null || email === "") continue;
    if (
      typeof email !== "string" ||
      typeof title !== "string" ||
      (loginName != null && typeof loginName !== "string") ||
      (rawId != null && typeof rawId !== "string" && typeof rawId !== "number")
    )
      throw new TypeError("The SharePoint people search response contains invalid user fields.");
    if (loginName?.includes("_adm") || loginName?.includes("adm_")) continue;

    const id = rawId == null || rawId === "" ? 0 : Number(rawId);
    if (!Number.isSafeInteger(id) || id < 0)
      throw new TypeError("The SharePoint people search response contains an invalid user ID.");
    users.push({ Title: title, LoginName: loginName || email, Id: id });
  }
  return users;
}

function normalizeSearchResults(value: unknown): SharepointChoiceSearchResult[] {
  if (!isSharepointChoiceSearchResults(value)) {
    throw new TypeError("The search response must be an array of result objects with numeric IDs.");
  }
  return value;
}

function isOutlookAttachmentResponse(value: unknown): value is OutlookAttachmentResponse {
  return (
    isRecord(value) && typeof value["contentBytes"] === "string" && typeof value["lastModifiedDateTime"] === "string"
  );
}

function isSharePointFileListItem(value: unknown): value is SharePointFileListItem {
  const file = isRecord(value) ? value["File"] : undefined;
  return (
    isRecord(file) &&
    typeof file["TimeCreated"] === "string" &&
    typeof file["TimeLastModified"] === "string" &&
    typeof file["ServerRelativeUrl"] === "string" &&
    typeof file["Name"] === "string"
  );
}

function readTransferData(transfer: DataTransfer, format: string): unknown {
  const serializedData = transfer.getData(format);
  return serializedData ? (JSON.parse(serializedData) as unknown) : undefined;
}

function attachmentContentBuffer(content: ArrayBuffer | Uint8Array | string): ArrayBuffer {
  if (content instanceof ArrayBuffer) return content;
  if (content instanceof Uint8Array) return Uint8Array.from(content).buffer;
  return new TextEncoder().encode(content).buffer;
}

@Component({
  selector: "app-choice",
  templateUrl: "./sharepoint-choice.component.html",
  styleUrls: ["./styles.scss"],
  standalone: true,
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [CommonModule, FormsModule, NgxEditorModule],
  providers: [
    {
      provide: ErrorHandler,
      useClass: SharepointChoiceLogging,
    },
  ],
})
export class SharepointChoiceComponent implements OnInit, OnDestroy {
  private static readonly maxFileSizeBytes = 512 * 1024 * 1024;
  private static readonly maxExtractionDepth = 8;

  private get maxFileSizeLabel(): string {
    return this.fileSize(SharepointChoiceComponent.maxFileSizeBytes);
  }

  private get officeApi(): typeof Office | undefined {
    return (window as SharepointChoiceWindow).Office;
  }

  // form object to bind to
  // not using field direct as simple objects are clones not references
  // not using ngModel as it requires extra boilerplate such as ngModel=form['field'] spec=spec['field'] etc for each field
  @Input() set form(value: SharepointChoiceForm) {
    if (!isRecord(value)) throw new TypeError("The choice form must be an object.");
    this._form = value;
    this.invalidateAttachmentCache();
    this.initializeField();
    void this.lookupLabel();
    this.chRef.markForCheck();
  }
  get form(): SharepointChoiceForm {
    return this._form;
  }
  private _form!: SharepointChoiceForm;

  // specification of the list fields
  // as spec must match datatype and form data for successful save then tie all of these together by this.field not separately
  @Input() set spec(value: SharepointChoiceList | undefined) {
    this._spec = value;
    this.initializeField();
    void this.lookupLabel();
    this.chRef.markForCheck();
  }
  get spec(): SharepointChoiceList | undefined {
    return this._spec;
  }
  private _spec?: SharepointChoiceList;

  // manually override any spec above
  // prefer send as string as passing object kills large form performance
  @Input() set override(value: string | SharepointChoiceField | undefined) {
    // if override passed in, make sure its an object else convert
    if (value) this._override = typeof value === "string" ? JSON.parse(value) : value;
    else this._override = undefined;
    this.initializeField();
    void this.lookupLabel();
    this.chRef.markForCheck();
  }
  get override(): SharepointChoiceField | undefined {
    return this._override;
  }
  private _override?: SharepointChoiceField;

  @Input() set disabled(value: boolean) {
    // get disabled state from outside
    this._disabled = value;
    this.chRef.markForCheck();
  }
  get disabled(): boolean {
    return this._disabled;
  }
  private _disabled: boolean = false;

  @Input() set versions(value: SharepointChoiceForm[] | undefined) {
    this._versions = value;
    this.chRef.markForCheck();
  }
  get versions(): SharepointChoiceForm[] | undefined {
    return this._versions;
  }
  private _versions?: SharepointChoiceForm[];

  @Input() set refreshKey(value: unknown) {
    this._refreshKey = value;
    this.invalidateAttachmentCache();
    this.initializeField(true);
    void this.lookupLabel();
    this.chRef.markForCheck();
  }
  get refreshKey(): unknown {
    return this._refreshKey;
  }
  private _refreshKey?: unknown;

  @Input() prefix: string = ""; // prefix name attributes for uniqness, usefull for nesting

  @Input() set field(value: string) {
    this._field = value;
    this.initializeField();
    void this.lookupLabel();
    this.chRef.markForCheck();
  }
  get field(): string {
    return this._field;
  }
  private _field!: string;

  // could also input hidden rather than elRef but that doesnt pick up non angular html hidden or parent hidden
  @Output() change = new EventEmitter<SharepointChoiceChangeEvent>(); // emit changes to parent through (change) binding

  @Input() set text(value: SharepointChoiceTextOptions | undefined) {
    this._text = value || {};
    this.initializeField();
    this.chRef.markForCheck();
  }
  get text(): SharepointChoiceTextOptions {
    return this._text;
  }
  private _text: SharepointChoiceTextOptions = {};

  @Input() set select(value: SharepointChoiceSelectOptions | undefined) {
    this._select = value || {};
    this.chRef.markForCheck();
  }
  get select(): SharepointChoiceSelectOptions {
    return this._select;
  }
  private _select: SharepointChoiceSelectOptions = {};

  @Input() set file(value: SharepointChoiceFileOptions | undefined) {
    this._file = value || {};
    this.invalidateAttachmentCache();
    this.generatedSpecs.clear();
    this.chRef.markForCheck();
  }
  get file(): SharepointChoiceFileOptions {
    return this._file;
  }
  private _file: SharepointChoiceFileOptions = {};

  tooltip?: boolean;

  editor?: Editor;
  // rtf menu items
  toolbar: Toolbar = [
    ["text_color", "background_color"],
    ["bold", "italic", "underline", "strike"],
    ["ordered_list", "bullet_list"],
    [{ heading: ["h1", "h2", "h3", "h4", "h5", "h6"] }],
    ["code", "blockquote"],
    ["link", "image"],
    ["align_left", "align_center", "align_right", "align_justify"],
  ];

  filesOver?: boolean;

  office: {
    // type of office addin
    type: string | null;
    // loading state for office ops
    loading: boolean;
    // errors to show
    alert?: string;
  } = { type: null, loading: false };

  // varies based on the search source therefore not explicitly defined
  results: SharepointChoiceSearchResult[] = [];
  pos: number = -1;
  sort?: string;
  otherText: string = "";
  lookupLabels: SharepointChoiceLookupLabel[] = [];

  users: SharepointChoiceUser[] = [];

  versionsDisplayed?: boolean;

  private textKey?: Subject<string>;
  private userKey?: Subject<string>;

  private loadingUserIds: number[] = [];
  private displayedUsers: SharepointChoiceUser[] = [];
  private userLoadRetryAt = new Map<number, number>();
  private destroyed = false;
  private textSearchRequestId = 0;
  private userSearchRequestId = 0;
  private lookupLabelRequestId = 0;
  private initializedForm?: SharepointChoiceForm;
  private initializedSpec?: SharepointChoiceList;
  private initializedOverride?: SharepointChoiceField;
  private initializedField?: string;
  private initializedType?: string;
  private attachmentCache?: {
    source: SharepointChoiceAttachment[];
    options?: SharepointChoiceFileOptions;
    archive?: string;
    doctype?: string;
    filter: string;
    sort: string;
    view: number;
    value: SharepointChoiceAttachment[];
  };
  private generatedSpecs = new Map<
    string,
    {
      fieldSpec?: SharepointChoiceField;
      doctypes?: string[];
      value: SharepointChoiceList;
    }
  >();

  // subscription to refresh bus
  private refreshSub?: Subscription;
  // unique instance id for this component
  private static nextInstanceId = 1;
  private readonly instanceId = SharepointChoiceComponent.nextInstanceId++;

  public constructor(
    private elRef: ElementRef,
    private chRef: ChangeDetectorRef,
    private refreshBus: SharepointChoiceRefresh,
    private errorHandler: ErrorHandler,
  ) {}

  // on init, destroy
  public ngOnInit(): void {
    this.elRef.nativeElement.setAttribute("field", this.field);
    this.initializeField();
    void this.lookupLabel();

    this.refreshSub = this.refreshBus.changes$.subscribe((event) => {
      if (event.sourceId === this.instanceId) return;
      if ((event.form && event.form === this._form) || (event.spec && event.spec === this._spec)) {
        this.invalidateAttachmentCache();
        this.initializeField(true);
        void this.lookupLabel();
        this.chRef.markForCheck();
      }
    });
  }

  public ngOnDestroy(): void {
    this.destroyed = true;
    this.textSearchRequestId++;
    this.userSearchRequestId++;
    this.lookupLabelRequestId++;
    this.editor?.destroy();
    this.textKey?.complete();
    this.userKey?.complete();
    this.refreshSub?.unsubscribe();
  }

  private initializeField(force: boolean = false): void {
    const form = this._form;
    const field = this._field;
    if (!form || !field) return;

    const type = this.get("TypeAsString");
    if (!type) return;

    const identityChanged =
      this.initializedForm !== form ||
      this.initializedSpec !== this._spec ||
      this.initializedOverride !== this._override ||
      this.initializedField !== field ||
      this.initializedType !== type;
    const changedInputs = force || identityChanged;

    if (changedInputs) {
      const hasDefaultableEmptyValue = form[field] === undefined && form[field + "Id"] === undefined;

      this.textSearchRequestId++;
      this.userSearchRequestId++;
      this.lookupLabelRequestId++;
      if (identityChanged) {
        this.results = [];
        this.users = [];
        this.lookupLabels = [];
      }

      if (type === "MultiChoice") {
        const value = form[field];
        if (!readResults(value)) {
          const record = isRecord(value) ? value : {};
          record["__metadata"] = { type: "Collection(Edm.String)" };
          record["results"] = Array.isArray(value) ? value : value == null ? [] : [value];
          form[field] = record;
        }
      } else if (type === "UserMulti" || type === "LookupMulti") {
        const idField = field + "Id";
        const value = form[idField];
        if (!readResults(value)) {
          const record = isRecord(value) ? value : {};
          record["__metadata"] = { type: "Collection(Edm.Int32)" };
          record["results"] = Array.isArray(value) ? value : value == null ? [] : [value];
          form[idField] = record;
        }
      } else if (type === "URL") {
        const value = form[field];
        const record = isRecord(value) ? value : undefined;
        if (!record || record["Description"] == null) form[field] = { Description: "", Url: value || "" };
      } else if (type === "Geolocation") {
        const value = form[field];
        const record = isRecord(value) ? value : undefined;
        if (!record || record["Latitude"] === undefined || record["Longitude"] === undefined)
          form[field] = { Latitude: "", Longitude: "" };
      } else if (type === "Attachments") {
        if (!readResults(form[field])) form[field] = { results: [] };
        this.officeAddin();
        if (!this.sort || this.sort === "-") this.sort = "Created";
      }

      if (hasDefaultableEmptyValue) {
        const value = this.get("DefaultValue");
        if (value !== undefined && value !== null && value !== "") {
          if (type === "MultiChoice") setResults(form, field, [value]);
          else if (type === "UserMulti" || type === "LookupMulti") setResults(form, field + "Id", [Number(value)]);
          else if (type === "User" || type === "Lookup") form[field + "Id"] = value;
          else if (type === "DateTime") form[field] = new Date(value);
          else if (type === "Integer") form[field] = parseInt(value);
          else if (type === "Number" || type === "Currency") form[field] = parseFloat(value);
          else if (type === "Boolean") form[field] = value === "1";
          else if (type === "URL") form[field] = { Description: "", Url: value };
          else form[field] = value;
        }
      }

      if (this.get("ReadOnlyField")) this.disabled = true;

      this.initializedForm = form;
      this.initializedSpec = this._spec;
      this.initializedOverride = this._override;
      this.initializedField = field;
      this.initializedType = type;
    }

    if (type === "Note" && this.get("RichText") && !this.editor)
      this.editor = new Editor({ plugins: [this.createClipboardImagePlugin()] });

    if ((type === "Lookup" || type === "LookupMulti" || (type === "Text" && this.text?.search)) && !this.textKey) {
      this.textKey = new Subject<string>();
      this.textKey.pipe(debounceTime(250)).subscribe((key) => this.onUpTextSearch(key));
    }
    if ((type === "User" || type === "UserMulti") && !this.userKey) {
      this.userKey = new Subject<string>();
      this.userKey.pipe(debounceTime(250)).subscribe((key) => this.onUpUserSearch(key));
    }
  }

  private fileToDataUrl(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (typeof reader.result === "string") resolve(reader.result);
        else reject(new Error("Pasted image did not produce a data URL."));
      };
      reader.onerror = () => reject(reader.error ?? new Error("Unable to read pasted image."));
      reader.onabort = () => reject(new Error("Reading pasted image was aborted."));
      reader.readAsDataURL(file);
    });
  }

  private createClipboardImagePlugin(): Plugin {
    return new Plugin({
      props: {
        handlePaste: (view, event, slice) => {
          const items = event.clipboardData?.items;
          if (!items) return false;

          const files = Array.from(items)
            .filter(
              (item) =>
                item.kind === "file" &&
                item.type.toLowerCase().startsWith("image/") &&
                item.type.toLowerCase() !== "image/svg+xml",
            )
            .map((item) => item.getAsFile())
            .filter((file): file is File => file !== null);
          if (!files.length) return false;

          const acceptableFiles = files.filter((file) => file.size <= MAX_PASTED_IMAGE_SIZE_BYTES);
          const oversizedCount = files.length - acceptableFiles.length;
          if (oversizedCount) {
            alert(
              `${oversizedCount === 1 ? "A pasted image was" : `${oversizedCount} pasted images were`} skipped because each image must be ${MAX_PASTED_IMAGE_SIZE_MB} MB or smaller.`,
            );
          }

          let hasImageSource = false;
          slice.content.descendants((node) => {
            if (node.type.name === "image" && typeof node.attrs["src"] === "string" && node.attrs["src"].trim())
              hasImageSource = true;
          });
          if (hasImageSource) {
            if (!oversizedCount) return false;
            event.preventDefault();
            return true;
          }

          event.preventDefault();
          if (!acceptableFiles.length) return true;

          void Promise.all(acceptableFiles.map((file) => this.fileToDataUrl(file)))
            .then((sources) => {
              if (this.destroyed || this.editor?.view !== view) return;

              const imageType = view.state.schema.nodes["image"];
              if (!imageType) throw new Error("The editor schema does not support images.");

              let transaction = view.state.tr;
              for (const src of sources)
                transaction = transaction.replaceSelectionWith(imageType.create({ src }), false);
              view.dispatch(transaction.scrollIntoView());
            })
            .catch((error: unknown) => this.errorHandler.handleError(error));

          return true;
        },
      },
    });
  }

  private invalidateAttachmentCache(): void {
    this.attachmentCache = undefined;
  }

  /*
  Common parts between multiple field types or minimal functions
  */

  // show or hide tooltips
  private showHideTooltip(show: boolean): void {
    this.tooltip = show;
    this.chRef.markForCheck();
  }

  // are there different field version values shown
  private versionsToggle(): string {
    if (!this.versionsDisplayed) {
      this.versionsDisplayed = true;
      this.chRef.markForCheck();
    }
    return "";
  }

  // show numbers using locale separators without forcing trailing zeros
  private niceNumber(bcp47: string | undefined = undefined): string {
    const value = this.form[this.field];
    if ((typeof value !== "number" && typeof value !== "string") || value === "") return "";
    const numericValue = Number(value);
    if (!Number.isFinite(numericValue)) return "";

    return new Intl.NumberFormat(bcp47, {
      useGrouping: true,
      maximumFractionDigits: 20,
    }).format(numericValue);
  }

  private numberSet(inputValue: string | undefined, bcp47: string | undefined = undefined): void {
    if (!inputValue) {
      this.form[this.field] = null;
      this.changed();
      return;
    }

    const formatParts = new Intl.NumberFormat(bcp47).formatToParts(-12345.6);
    const decimalSeparator = formatParts.find((part) => part.type == "decimal")?.value || ".";
    const groupSeparator = formatParts.find((part) => part.type == "group")?.value || ",";
    const minusSign = formatParts.find((part) => part.type == "minusSign")?.value || "-";
    const localizedDigits = new Map(
      new Intl.NumberFormat(bcp47, { useGrouping: false })
        .format(9876543210)
        .split("")
        .reverse()
        .map((digit, index) => [digit, index.toString()]),
    );

    let normalizedValue = inputValue
      .trim()
      .replace(/[\s\u00A0\u202F]/g, "")
      .split("")
      .map((character) => localizedDigits.get(character) || character)
      .join("");

    const escapeRegex = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    normalizedValue = normalizedValue
      .replace(new RegExp(escapeRegex(groupSeparator), "g"), "")
      .replace(new RegExp(escapeRegex(decimalSeparator), "g"), ".")
      .replace(new RegExp(escapeRegex(minusSign), "g"), "-");

    const negativeSign = normalizedValue.startsWith("-") ? "-" : "";
    normalizedValue = normalizedValue.replace(/-/g, "");

    const decimalPosition = normalizedValue.indexOf(".");
    if (decimalPosition >= 0)
      normalizedValue =
        normalizedValue.slice(0, decimalPosition + 1) + normalizedValue.slice(decimalPosition + 1).replace(/\./g, "");

    normalizedValue = (negativeSign + normalizedValue).replace(/[^0-9.-]/g, "");
    let numericValue = parseFloat(normalizedValue);
    if (isNaN(numericValue)) {
      this.form[this.field] = null;
      this.changed();
      return;
    }
    const minimum = this.get("Min");
    if (minimum != null && numericValue < Number(minimum)) numericValue = Number(minimum);
    const maximum = this.get("Max");
    if (maximum != null && numericValue > Number(maximum)) numericValue = Number(maximum);
    this.form[this.field] = numericValue;

    this.changed();
  }

  private dateSet(value: Date | string | null): void {
    this.form[this.field] = value;

    this.changed();
  }

  private lcidToBcp47(lcid: number): string {
    return LCID_TO_BCP47[lcid] || "en-US";
  }

  private symbol(bcp47: string): string {
    const regionCode = bcp47.split("-").findLast((part) => /^[A-Z]{2}$/.test(part));
    const currency = regionCode ? CURRENCY_BY_REGION[regionCode] || "USD" : "USD";
    const formatter = new Intl.NumberFormat(bcp47, {
      style: "currency",
      currency,
      currencyDisplay: "narrowSymbol",
    });
    return formatter.formatToParts(0).find((part) => part.type == "currency")?.value || currency;
  }

  private pattern(bcp47: string | undefined = undefined): string {
    if (new Intl.NumberFormat(bcp47).formatToParts(12345.6).find((part) => part.type === "decimal")?.value === ",")
      return "-{0,1}[0-9\. ']*,{0,1}[0-9\. ']*";
    return "-{0,1}[0-9, ']*\.{0,1}[0-9, ']*";
  }

  // field required based on spec but required is not needed for hidden/disabled items
  private required(): boolean {
    if (this.disabled || this.isEffectivelyHidden() || !this.get("Required")) return false;
    return true;
  }

  private isEffectivelyHidden(): boolean {
    let currentElement: HTMLElement | null = this.elRef.nativeElement;
    while (currentElement) {
      if (currentElement.hidden) return true;
      const style = window.getComputedStyle(currentElement);
      if (style.display === "none" || style.visibility === "hidden" || style.visibility === "collapse") return true;
      currentElement = currentElement.parentElement;
    }
    return false;
  }

  // get outcomes of non standard fields into a plain text field for [required] to be triggered automatically
  private validator(type?: string): string {
    switch (type) {
      case "User":
      case "Lookup":
        return this.form[this.field + "Id"] ? "true" : "";
      case "UserMulti":
      case "LookupMulti":
        return (readResults(this.form[this.field + "Id"])?.length ?? 0) > 0 ? "true" : "";
      case "Attachments":
        return this.attachments().length > 0 ? "true" : "";
      default:
        return this.form[this.field] ? "true" : "";
    }
  }

  private hasGeolocationValue(): boolean {
    const locationValue = this.form[this.field];
    const location = isRecord(locationValue) ? locationValue : undefined;
    if (!location) return false;

    return (
      location["Latitude"] !== undefined &&
      location["Latitude"] !== null &&
      location["Latitude"] !== "" &&
      location["Longitude"] !== undefined &&
      location["Longitude"] !== null &&
      location["Longitude"] !== ""
    );
  }

  private geolocationHref(): string {
    const locationValue = this.form[this.field];
    const location = isRecord(locationValue) ? locationValue : undefined;
    return `https://www.bing.com/maps/sharing?cp=${location?.["Latitude"] ?? ""}%7E${location?.["Longitude"] ?? ""}&lvl=17.5&webglerror=a`;
  }

  // max length character countdown
  private remaining(max?: number): number {
    const maxLength = this.get("MaxLength");
    // no max limit in field spec or because of field type then use a number bigger than any remaining() >= ... values
    if (!maxLength && !max) return 255;
    const value = this.form[this.field];
    return (maxLength ?? max ?? 255) - (typeof value === "string" ? value.length : 0);
  }

  private templateResults(field: string): unknown[] {
    return readResults(this.form[field]) ?? [];
  }

  private templateDateValue(field: string): string | number | Date | null {
    const value = this.form[field];
    return typeof value === "string" || typeof value === "number" || value instanceof Date ? value : null;
  }

  private templateObjectString(field: string, property: string): string {
    const value = isRecord(this.form[field]) ? this.form[field][property] : undefined;
    return typeof value === "string" || typeof value === "number" ? String(value) : "";
  }

  private setTemplateObjectString(field: string, property: string, value: string): void {
    const record = isRecord(this.form[field]) ? this.form[field] : {};
    record[property] = value;
    this.form[field] = record;
  }

  private templateHasValue(form: object, field: string): boolean {
    return isRecord(form) && Boolean(form[field]);
  }

  private templateValue(field: string): string | number | boolean {
    const value = this.form[field];
    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
    if (value instanceof Date) return value.toLocaleString();
    const results = readResults(value);
    return results ? results.map((item) => String(item)).join(", ") : "";
  }

  private templateUserId(field: string): number | undefined {
    const value = this.form[field];
    if (typeof value === "number" && Number.isSafeInteger(value)) return value;
    if (typeof value === "string" && value.trim()) {
      const parsed = Number(value);
      if (Number.isSafeInteger(parsed)) return parsed;
    }
    return undefined;
  }

  private templateUserIds(field: string): number[] {
    return (readResults(this.form[field]) ?? []).filter(
      (value): value is number => typeof value === "number" && Number.isSafeInteger(value),
    );
  }

  private templateVersionDate(version: SharepointChoiceForm): string | number | Date | null {
    const value = version["Created"];
    return typeof value === "string" || typeof value === "number" || value instanceof Date ? value : null;
  }

  private templateVersionEditor(version: SharepointChoiceForm): string {
    const editor = isRecord(version["Editor"]) ? version["Editor"] : undefined;
    const value = editor?.["LookupValue"];
    return typeof value === "string" || typeof value === "number" ? String(value) : "";
  }

  private templateVersionChanged(
    current: SharepointChoiceForm,
    previous: SharepointChoiceForm,
    field: string,
  ): boolean {
    const value = current[field];
    return value !== previous[field] && (Boolean(value) || (value != null && String(value).length > 0));
  }

  // gets the required field properties and/or any overrides to determine which field type etc to display
  private get(propertyName: "TypeAsString"): SharepointChoiceField["TypeAsString"];
  private get(propertyName: "Choices"): string[];
  private get(propertyName: "Required" | "ReadOnlyField" | "RichText" | "AppendOnly"): boolean | undefined;
  private get(propertyName: "MaxLength" | "DisplayFormat" | "LCID" | "SelectionGroup"): number | undefined;
  private get(propertyName: "Min" | "Max"): number | string | undefined;
  private get(
    propertyName:
      | "DefaultValue"
      | "InternalName"
      | "Title"
      | "Description"
      | "Format"
      | "FillInChoice"
      | "LookupList"
      | "LookupField",
  ): string | undefined;
  private get(propertyName: string): unknown;
  private get(propertyName: string): unknown {
    let value: unknown = this.override ? Reflect.get(this.override, propertyName) : undefined;
    // if no override then get from field spec selected
    if (value === undefined || value === null) {
      const fieldSpec = (this.spec || {})[this.field.replace(/^OData_/, "")] ?? (this.spec || {})[this.field];
      if (fieldSpec) value = Reflect.get(fieldSpec, propertyName);
    }

    if (propertyName == "Choices" && value == null) return [];
    if (propertyName == "Title" && value == null) return this.friendlyName(this.field);
    if ((propertyName == "Min" || propertyName == "Max") && value == null && this.get("TypeAsString") == "DateTime")
      return (
        (propertyName == "Min" ? "1970-01-01" : "9999-12-31") + (this.get("DisplayFormat") == 1 ? "T00:00:00" : "")
      );

    const results =
      value !== null && typeof value === "object" && "results" in value ? Reflect.get(value, "results") : undefined;
    return results ?? value;
  }

  // any field changes trigger for relevant updates
  private changed(): void {
    const fieldType = this.get("TypeAsString");
    this.invalidateAttachmentCache();
    const changedField =
      fieldType == "Lookup" || fieldType == "LookupMulti" || fieldType == "User" || fieldType == "UserMulti"
        ? this.field + "Id"
        : this.field;
    const changedValue = readResults(this.form[changedField]) ?? this.form[changedField];

    this.change.emit({
      field: changedField,
      value: changedValue,
      target: this.elRef.nativeElement,
    });

    // Notify other instances (with OnPush) that share the same form/spec reference.
    this.refreshBus.emit({
      sourceId: this.instanceId,
      form: this._form,
      spec: this._spec,
      field: this.field,
    });

    queueMicrotask(() => {
      try {
        this.chRef.detectChanges();
      } catch {
        this.chRef.markForCheck();
      }
    });
  }

  // on multi select, not using ctrl key
  private changedM(event: Event): void {
    event.stopPropagation();
    const selectedOption = event.target as HTMLOptionElement;

    let scrollTop = 0;
    if (selectedOption.parentElement) scrollTop = selectedOption.parentElement.scrollTop;

    // use inner text as value gets adjusted by angular
    const selectedValue = selectedOption.innerText;

    let results = readResults(this.form[this.field]);
    if (!results) {
      const currentValue = this.form[this.field];
      results = Array.isArray(currentValue) ? currentValue : currentValue == null ? [] : [currentValue];
      const record = isRecord(currentValue) ? currentValue : {};
      record["__metadata"] = { type: "Collection(Edm.String)" };
      record["results"] = results;
      this.form[this.field] = record;
    }

    // if there are selected results set the field to add/remove the most recent click
    const selectedIndex = results.indexOf(selectedValue);
    if (selectedIndex >= 0) results.splice(selectedIndex, 1);
    else results.push(selectedValue);

    if (selectedOption.parentElement) {
      // manually toggle html select state as mousedown false prevents the toggle as well as prevent the unselect all
      selectedOption.selected = !selectedOption.selected;
      setTimeout(function () {
        selectedOption.parentElement ? (selectedOption.parentElement.scrollTop = scrollTop) : null;
      }, 10);
      setTimeout(function () {
        selectedOption.parentElement ? selectedOption.parentElement.focus() : null;
      }, 10);
    }

    // rebuild the results to ensure data triggered and display updated
    setResults(this.form, this.field, [...new Set(results)]);

    // push the change to any emitters
    this.changed();
  }

  // append only changes needs 1 way bind to form
  private changedA(event: Event): void {
    this.form[this.field] = event.target ? (event.target as HTMLInputElement).value : event;

    this.changed();
  }

  /*
  Common parts between text fields
  */

  private async onUpText(pressedKey: string | undefined): Promise<void> {
    if (!this.textKey) return;

    if (pressedKey == "ArrowDown") {
      if (this.pos < this.results.length - 1) this.pos++;
    } else if (pressedKey == "ArrowUp") {
      if (this.pos > 0) this.pos--;
    } else if (pressedKey == "Enter") {
      this.textSearchRequestId++;
      await this.selectedText(this.results[this.pos] ?? null);
    } else {
      if (pressedKey !== undefined) this.textSearchRequestId++;
      const fieldType = this.get("TypeAsString");
      const searchValue: unknown =
        fieldType == "Lookup" || fieldType == "LookupMulti" ? this.otherText.trim() : this.form[this.field];

      if (pressedKey !== undefined) {
        this.results = [];
        this.pos = -1;
      }
      if (typeof searchValue === "string" && searchValue) this.textKey.next(searchValue);
      else this.results = [];
    }
    this.chRef.markForCheck();
  }

  private async onUpTextSearch(searchText: string): Promise<void> {
    const requestId = ++this.textSearchRequestId;
    try {
      const fieldType = this.get("TypeAsString");
      let matches: SharepointChoiceSearchResult[] = [];

      if (fieldType == "Lookup" || fieldType == "LookupMulti") {
        const sharepointUtils = new SharepointChoiceUtils();
        const lookupListId = this.get("LookupList");
        const lookupFieldName = this.get("LookupField");

        if (lookupListId && lookupFieldName && searchText.trim()) {
          const escapedSearchText = searchText.replace(/'/g, "''");
          const response: unknown = await sharepointUtils.sp.web.lists
            .getById(lookupListId)
            .items.select("Id", lookupFieldName)
            .filter(`${lookupFieldName} ne null and substringof('${escapedSearchText}',${lookupFieldName})`)
            .top(10)();
          matches = normalizeSearchResults(response);
        }
      } else if (fieldType == "Text" && this.text?.search) {
        matches = normalizeSearchResults(await this.text.search(searchText, this.text.parent));
      }

      if (requestId !== this.textSearchRequestId) return;
      this.results = matches;
      this.pos = -1;
    } catch (error) {
      if (requestId !== this.textSearchRequestId) return;
      this.results = [];
      this.pos = -1;
      this.errorHandler.handleError(error);
    } finally {
      if (requestId === this.textSearchRequestId) this.chRef.markForCheck();
    }
  }

  private async selectedText(selectedResult: SharepointChoiceSearchResult | null): Promise<void> {
    this.textSearchRequestId++;
    const fieldType = this.get("TypeAsString");
    const isLookup = fieldType == "Lookup" || fieldType == "LookupMulti";
    const isMultiLookup = fieldType == "LookupMulti";
    if (!isLookup && !this.text?.search) return;

    if (!selectedResult) {
      if (isMultiLookup) {
        this.otherText = "";
        this.results = [];
        this.chRef.markForCheck();
        return;
      }
      if (isLookup) this.setLookupValue(null, "");
      else this.form[this.field] = null;
    } else if (isLookup) {
      const lookupFieldName = this.get("LookupField") ?? "";
      this.setLookupValue(selectedResult.Id, String(selectedResult[lookupFieldName] ?? ""));
    } else {
      this.form[this.field] = selectedResult[this.field];
      if (this.text.select) await this.text.select(selectedResult, this.text.parent);
    }

    this.results = [];
    this.changed();
  }

  private async lookupLabel(): Promise<void> {
    const requestId = ++this.lookupLabelRequestId;
    const type = this.get("TypeAsString");
    if (!this._form || !this.field || (type != "Lookup" && type != "LookupMulti")) return;

    const lookupListId = this.get("LookupList");
    const lookupFieldName = this.get("LookupField");
    if (!lookupListId || !lookupFieldName) {
      if (type == "Lookup") this.otherText = "";
      else this.lookupLabels = [];
      return;
    }

    const form = this.form;
    const idValue = form[this.field + "Id"];
    const rawLookupIds = type == "LookupMulti" ? (readResults(idValue) ?? []) : idValue == null ? [] : [idValue];
    const lookupIds = rawLookupIds
      .map((value) => (typeof value === "number" ? value : typeof value === "string" ? Number(value) : Number.NaN))
      .filter((value) => Number.isSafeInteger(value));

    if (type == "Lookup" && !lookupIds.length) {
      this.otherText = "";
      return;
    }
    if (type == "LookupMulti" && !lookupIds.length) {
      this.lookupLabels = [];
      return;
    }

    const sharepointUtils = new SharepointChoiceUtils();
    try {
      const labels = await Promise.all(
        lookupIds.map(async (id) => {
          const response: unknown = await sharepointUtils.sp.web.lists
            .getById(lookupListId)
            .items.getById(id)
            .select(lookupFieldName)();
          if (!isRecord(response)) throw new TypeError("The SharePoint lookup response must be an object.");
          const value = response[lookupFieldName];
          return { id, label: value == null ? "" : String(value) };
        }),
      );
      if (requestId !== this.lookupLabelRequestId || form !== this.form) return;
      const currentValue = form[this.field + "Id"];
      const rawCurrentIds =
        type == "LookupMulti" ? (readResults(currentValue) ?? []) : currentValue == null ? [] : [currentValue];
      const currentIds = rawCurrentIds
        .map((value) => (typeof value === "number" ? value : typeof value === "string" ? Number(value) : Number.NaN))
        .filter((value) => Number.isSafeInteger(value));
      if (currentIds.length !== lookupIds.length || currentIds.some((id, index) => id !== lookupIds[index])) return;

      if (type == "Lookup") this.otherText = labels[0]?.label || "";
      else this.lookupLabels = labels;
      this.chRef.markForCheck();
    } catch (error) {
      if (requestId !== this.lookupLabelRequestId) return;
      this.errorHandler.handleError(error);
    }
  }

  private setLookupValue(lookupId: number | null | undefined, lookupLabel: string): void {
    if (this.get("TypeAsString") == "LookupMulti") {
      if (lookupId == null) return;
      const results = readResults(this.form[this.field + "Id"]) ?? [];
      if (!results.includes(lookupId)) {
        results.push(lookupId);
        setResults(this.form, this.field + "Id", results);
        this.lookupLabels = [...this.lookupLabels, { id: lookupId, label: lookupLabel }];
      }
      this.otherText = "";
      return;
    }

    this.otherText = lookupLabel;
    this.form[this.field + "Id"] = lookupId;
  }

  private setLookupSearchText(value: string): void {
    this.otherText = value;
    this.textSearchRequestId++;
    this.results = [];
    this.pos = -1;
    this.chRef.markForCheck();
  }

  private removeLookup(lookupId: number): void {
    const idField = this.form[this.field + "Id"];
    const results = readResults(idField);
    if (!results) return;
    setResults(
      this.form,
      this.field + "Id",
      results.filter((value) => value !== lookupId),
    );
    this.lookupLabels = this.lookupLabels.filter((item) => item.id !== lookupId);
    this.changed();
  }

  /*
  Common parts between choice fields
  */

  // many choices render bigger box
  private multiLargeorSmall(choices: string[] = this.choices()): string {
    return choices.length > 10 ? "multilarge" : "multismall";
  }

  // choices need filtering
  private choices(): string[] {
    const choices = this.get("Choices");
    // use any provided filter, pass in the form for context or specific filters but not the remainder of this object
    const filterChoices = this.select?.filter;
    const visibleChoices = filterChoices
      ? choices.filter((choice, index, allChoices) => filterChoices(choice, index, allChoices, { form: this.form }))
      : choices;
    // common filters
    return visibleChoices.filter((choice) => {
      // exclude unselected items on disabled fields
      if (
        this.disabled &&
        this.form[this.field] &&
        readResults(this.form[this.field]) &&
        !readResults(this.form[this.field])!.includes(choice)
      )
        return false;
      // filter on search above multichoice field
      if (this.otherText && !choice.toLowerCase().includes(this.otherText.toLowerCase())) return false;
      // else true
      return true;
    });
  }

  // selected field option not in available choices, i.e. other
  private inChoices(choices: string[]): boolean {
    const value = this.form[this.field];
    if (!value) return true;
    return typeof value === "string" && choices.includes(value);
  }

  // on single selection change
  private selChangeS(value: string): void {
    this.form[this.field] = value;

    this.changed();
  }

  /*
  Common parts between file upload field types
  */

  // files post filtering
  private attachments(): SharepointChoiceAttachment[] {
    const source = this.attachmentResults();

    const options = this.file;
    const archive = options?.archive;
    const doctype = options?.doctype;
    const filter = this.otherText;
    const sort = this.sort || "Created";
    const view = options?.view || 0;
    const cached = this.attachmentCache;
    if (
      cached &&
      cached.source === source &&
      cached.options === options &&
      cached.archive === archive &&
      cached.doctype === doctype &&
      cached.filter === filter &&
      cached.sort === sort &&
      cached.view === view
    )
      return cached.value;

    const value = source
      .filter((attachment) => {
        if (view == 0 || !archive || attachment.ListItemAllFields == null) return true;
        if (view == 1 && !attachment.ListItemAllFields[archive]) return true;
        if (view == -1 && attachment.ListItemAllFields[archive]) return true;
        return false;
      })
      .filter((attachment) => {
        if (!filter || !doctype) return true;
        if (!attachment.ListItemAllFields || !attachment.ListItemAllFields[doctype]) return true;
        if (attachment.ListItemAllFields[doctype] == filter) return true;
        return false;
      })
      .sort((a, b) => {
        const key = sort.replace(/^[-\+]/, "");
        const direction = sort.startsWith("-") ? -1 : 1;
        const aValue = a.ListItemAllFields?.[key];
        const bValue = b.ListItemAllFields?.[key];
        if (aValue == null) return bValue == null ? 0 : 1;
        if (bValue == null) return -1;
        if (aValue < bValue) return direction;
        if (aValue > bValue) return -direction;
        return 0;
      });

    this.attachmentCache = {
      source,
      options,
      archive,
      doctype,
      filter,
      sort,
      view,
      value,
    };
    return value;
  }

  private attachmentResults(): SharepointChoiceAttachment[] {
    const results = readResults(this.form[this.field]);
    if (!results) return [];
    if (!results.every(isSharepointChoiceAttachment)) {
      throw new TypeError(`Invalid attachment data for SharePoint field ${this.field}.`);
    }
    return results;
  }

  private enterKey(event: Event): void {
    event.preventDefault();
  }

  private friendlyName(name?: string): string {
    if (!name) return "";
    return name.replace(/_x0020_/g, " ").replace(/([a-z])([A-Z])/g, "$1 $2");
  }

  private hasChecked(): boolean {
    // check if there are any files that have been checked
    return this.attachmentResults().some((attachment) => attachment.Checked);
  }

  private setClasses(event: Event): void {
    const selectElement = event.target as HTMLSelectElement;
    // set the class of the file
    this.attachmentResults().forEach((attachment) => {
      if (attachment.Checked) {
        if (attachment.ListItemAllFields && this.file?.doctype)
          attachment.ListItemAllFields[this.file.doctype] = selectElement?.value;
        attachment.Checked = false;
      }
    });
    if (selectElement) selectElement.value = "";

    this.changed();
  }

  private changeSort(field?: string): void {
    if (!field) return;
    if (this.sort == "-" + field) {
      this.sort = "+" + field;
    } else if (this.sort == "+" + field) {
      this.sort = "";
    } else {
      this.sort = "-" + field;
    }
    this.invalidateAttachmentCache();
  }

  private additionalKeys(fieldSpecs?: SharepointChoiceList): Array<string> {
    // get keys of object
    if (!fieldSpecs || typeof fieldSpecs != "object" || this.field == "Attachments") return [];
    return Object.keys(fieldSpecs).filter(
      (fieldName) => fieldName != this.file?.doctype && fieldName != this.file?.notes,
    );
  }

  private loadOrGenerateSpec(field?: string, type?: SharepointChoiceField["TypeAsString"]): SharepointChoiceList {
    if (!field) return {};

    const key = `${field}\u0000${type || ""}`;
    const fieldSpec = this.file?.spec?.[field];
    const doctypes = this.file?.doctypes;
    const cached = this.generatedSpecs.get(key);
    if (cached && cached.fieldSpec === fieldSpec && cached.doctypes === doctypes) return cached.value;

    const value: SharepointChoiceList = {
      [field]: fieldSpec || {
        TypeAsString: type === "Choice" ? "Choice" : "Text",
        InternalName: field,
        Title: "",
        Choices: doctypes || [],
      },
    };
    this.generatedSpecs.set(key, { fieldSpec, doctypes, value });
    return value;
  }

  private usedTypes(): Array<string> {
    const configuredTypes = this.file?.doctypes || [];
    if (!this.file?.doctype || !readResults(this.form[this.field])) return [...new Set(configuredTypes)];

    const doctypeFieldName = this.file.doctype;
    const attachments = this.attachmentResults();
    const attachmentTypes = new Set<string>();
    for (const attachment of attachments) {
      const fieldValues = attachment.ListItemAllFields;
      if (!fieldValues || !(doctypeFieldName in fieldValues)) continue;
      const documentType: unknown = fieldValues[doctypeFieldName];
      if (typeof documentType === "string" && documentType) attachmentTypes.add(documentType);
    }
    return [...attachmentTypes].sort();
  }

  private newTab(attachment: SharepointChoiceAttachment, _event: Event): void {
    window.open(`${attachment.ServerRelativeUrl || "#"}?Web=1`, "_blank");
  }

  private width(): string {
    let maxDoctypeLength = 0;
    if (this.file?.doctypes)
      for (let index = 0; index < this.file.doctypes.length; index++) {
        const doctype = this.file.doctypes[index];
        if (!doctype) continue;

        const doctypeLength = doctype.length;
        if (doctypeLength > maxDoctypeLength) maxDoctypeLength = doctypeLength;
      }
    if (maxDoctypeLength == 0) return "";
    return `width: ${maxDoctypeLength}ch`;
  }

  private fileSize(fileLength?: number | null): string {
    if (fileLength == null || !isFinite(fileLength) || fileLength < 0) return "";

    const units = ["B", "KB", "MB", "GB", "TB"];
    let formattedValue = fileLength;
    let unitIndex = 0;

    while (formattedValue >= 1024 && unitIndex < units.length - 1) {
      formattedValue /= 1024;
      unitIndex++;
    }

    const decimalPlaces = formattedValue >= 10 || unitIndex == 0 ? 0 : 1;
    return `${formattedValue.toFixed(decimalPlaces).replace(/\.0$/, "")} ${units[unitIndex]}`;
  }

  private async delete(attachment?: SharepointChoiceAttachment, deletePermanently: boolean = false) {
    if (!attachment) return;
    if (!attachment.ServerRelativeUrl) {
      // not uploaded then exclude from potential upload
      setResults(
        this.form,
        this.field,
        this.attachmentResults().filter(
          (existingAttachment: SharepointChoiceAttachment) => existingAttachment.FileName != attachment.FileName,
        ),
      );
    } else {
      // if uploaded already
      if (!this.file?.archive || this.field == "Attachments") {
        // no archive flag or its attachments so no archiving, then flag for deletion
        if (deletePermanently && !window.confirm("Are you sure you wish to delete this file?")) return;
        attachment.Deleted = deletePermanently;
      } else {
        // toggle archived or delete flag
        if (attachment.ListItemAllFields?.[this.file.archive] && (deletePermanently || attachment.Deleted)) {
          if (deletePermanently && !window.confirm("Are you sure you wish to delete this file?")) return;
          attachment.Deleted = deletePermanently;
        } else if (attachment.ListItemAllFields) attachment.ListItemAllFields[this.file.archive] = deletePermanently;
      }
    }

    this.changed();
  }

  // add attachment to array
  private async add(inputFile: HTMLInputElement | DataTransfer) {
    if (!inputFile.files?.length) return;

    const filesToAdd = Array.from(inputFile.files);
    const errors: string[] = [];

    const readAndAppendFile = (fileToAdd: File): Promise<void> => {
      return new Promise((resolve, reject) => {
        const fileReader = new FileReader();

        fileReader.onload = () => {
          if (!(fileReader.result instanceof ArrayBuffer)) {
            reject("Expected FileReader to return an ArrayBuffer");
            return;
          }

          this.appendFile(fileToAdd.name, fileReader.result, this.attachmentResults())
            .then(resolve)
            .catch((error) => {
              reject(error ?? "Unknown file add error");
            });
        };

        fileReader.onerror = () => {
          reject(fileReader.error ?? "Unknown file read error");
        };

        fileReader.onabort = () => {
          reject("File read was aborted");
        };

        fileReader.readAsArrayBuffer(fileToAdd);
      });
    };

    for (const fileToAdd of filesToAdd) {
      if (fileToAdd.size > SharepointChoiceComponent.maxFileSizeBytes) {
        errors.push(`File error: ${fileToAdd.name} exceeds the ${this.maxFileSizeLabel} file size limit`);
        continue;
      }

      try {
        await readAndAppendFile(fileToAdd);
      } catch (error: unknown) {
        errors.push(`File error: ${fileToAdd.name} with error ${MapErrors(error ?? "Unknown error")}`);
      }
    }

    if (inputFile instanceof HTMLInputElement) {
      setTimeout(() => {
        inputFile.value = "";
      }, 0);
    }

    if (errors.length > 0) {
      const errorMessages = errors.join("\n");
      alert(`Error adding file(s):\n\n${errorMessages}`);
      throw errorMessages;
    }
  }

  // gets a drag and drop new outlook item which includes ids not file data and adds to the files array
  private async outlook(transfer: DataTransfer): Promise<void> {
    const sharepointUtils = new SharepointChoiceUtils();
    const errors: string[] = [];

    // one or more email messages dropped
    const messageListData =
      readTransferData(transfer, "multimaillistmessagerows") || readTransferData(transfer, "maillistrow");
    if (messageListData) {
      if (typeof messageListData !== "object" || messageListData === null)
        throw new Error("The dropped Outlook message list has an invalid format.");

      const rawMailboxInfos = Reflect.get(messageListData, "mailboxInfos");
      const rawSubjects = Reflect.get(messageListData, "subjects");
      const rawLatestItemIds = Reflect.get(messageListData, "latestItemIds");
      if (
        !Array.isArray(rawMailboxInfos) ||
        !rawMailboxInfos.every(
          (mailboxInfo: unknown) =>
            typeof mailboxInfo === "object" &&
            mailboxInfo !== null &&
            typeof Reflect.get(mailboxInfo, "mailboxSmtpAddress") === "string",
        ) ||
        !Array.isArray(rawSubjects) ||
        !rawSubjects.every((subject: unknown) => typeof subject === "string") ||
        !Array.isArray(rawLatestItemIds) ||
        !rawLatestItemIds.every((messageId: unknown) => typeof messageId === "string")
      )
        throw new Error("The dropped Outlook message list has an invalid format.");

      const mailboxInfos = rawMailboxInfos as Array<{ mailboxSmtpAddress: string }>;
      const subjects = rawSubjects as string[];
      const latestItemIds = rawLatestItemIds as string[];
      for (const [messageIndex, mailbox] of mailboxInfos.entries()) {
        const subject = subjects[messageIndex];
        const messageId = latestItemIds[messageIndex];
        if (subject === undefined || messageId === undefined)
          throw new Error("The dropped Outlook message list is missing message details.");
        const fileName = `${subject.trim()}.eml`;
        try {
          // must double url encode any / in the message id
          const fileContent = await sharepointUtils.callApi(
            undefined,
            undefined,
            undefined,
            `https://graph.microsoft.com/v1.0/users/${mailbox.mailboxSmtpAddress}/messages/${messageId.replace(/\//g, "%252F")}/$value`,
            "GET",
            undefined,
            "text",
          );
          if (fileContent === null) throw new Error("The email request returned no message content.");

          await this.appendFile(
            fileName,
            new TextEncoder().encode(fileContent).buffer as ArrayBuffer,
            this.attachmentResults(),
          );
        } catch (error: unknown) {
          errors.push(`Email append error: ${fileName} with error ${MapErrors(error)}`);
        }
      }
    }

    // an attachment dropped
    const attachmentValue = readTransferData(transfer, "attachment");
    if (attachmentValue) {
      if (typeof attachmentValue !== "object" || attachmentValue === null)
        throw new Error("The dropped Outlook attachment has an invalid format.");

      const attachmentFile = Reflect.get(attachmentValue, "attachmentFile");
      const mailboxInfo = Reflect.get(attachmentValue, "mailboxInfo");
      if (
        typeof attachmentFile !== "object" ||
        attachmentFile === null ||
        typeof mailboxInfo !== "object" ||
        mailboxInfo === null
      )
        throw new Error("The dropped Outlook attachment has an invalid format.");
      const fileName = Reflect.get(attachmentFile, "name");
      const attachmentItemId = Reflect.get(attachmentFile, "attachmentItemId");
      const mailboxSmtpAddress = Reflect.get(mailboxInfo, "mailboxSmtpAddress");
      if (
        typeof fileName !== "string" ||
        typeof attachmentItemId !== "string" ||
        typeof mailboxSmtpAddress !== "string"
      )
        throw new Error("The dropped Outlook attachment has an invalid format.");

      // truncate the attachment id 28 chars and = from the end to get the parent message id
      const parentMessageId = attachmentItemId.substring(0, attachmentItemId.length - 29) + "=";

      try {
        // must double url encode any / in the message id and attachment id
        const attachmentResponse: unknown = await sharepointUtils.callApi(
          undefined,
          undefined,
          undefined,
          `https://graph.microsoft.com/v1.0/users/${mailboxSmtpAddress}/messages/${parentMessageId.replace(/\//g, "%252F")}/attachments/${attachmentItemId.replace(/\//g, "%252F")}`,
        );
        if (!isOutlookAttachmentResponse(attachmentResponse))
          throw new TypeError("The attachment response has an invalid format.");

        await this.appendFile(
          fileName,
          Uint8Array.from(atob(attachmentResponse.contentBytes), (character) => character.charCodeAt(0)).buffer,
          this.attachmentResults(),
          `Sent: ${new Date(attachmentResponse.lastModifiedDateTime)}`,
        );
      } catch (error: unknown) {
        errors.push(`Attachment append error: ${fileName} with error ${MapErrors(error)}`);
      }
    }

    // a teams file drop, only works for teams libraries not onedrive/chat
    const fileDropData = readTransferData(transfer, "application/x-item-keys");
    if (fileDropData) {
      const rawItemKeys =
        typeof fileDropData === "object" && fileDropData !== null ? Reflect.get(fileDropData, "itemKeys") : undefined;
      if (!Array.isArray(rawItemKeys) || !rawItemKeys.every((itemKey: unknown) => typeof itemKey === "string"))
        throw new Error("The dropped SharePoint file list has an invalid format.");

      const itemKeys = rawItemKeys as string[];
      for (const serializedItemKey of itemKeys) {
        let itemKey: SharePointFileItemKey | undefined;
        try {
          // the inner is still JSON encoded from teams
          const parsedItemKey: unknown = JSON.parse(serializedItemKey) as unknown;
          if (
            !Array.isArray(parsedItemKey) ||
            typeof parsedItemKey[1] !== "string" ||
            typeof parsedItemKey[2] !== "string" ||
            (typeof parsedItemKey[3] !== "string" && typeof parsedItemKey[3] !== "number")
          )
            throw new Error("The dropped SharePoint file reference has an invalid format.");
          itemKey = parsedItemKey as SharePointFileItemKey;

          const web = Web([sharepointUtils.sp.web, itemKey[1]]);
          const properties: unknown = await web
            .getFolderByServerRelativePath(itemKey[2].substring(itemKey[2].indexOf("/", 9)))
            .properties();
          if (!isRecord(properties))
            throw new TypeError("The dropped SharePoint folder properties have an invalid format.");
          const folderProperties: SharePointFolderProperties = properties;

          const listTitle =
            folderProperties["vti_x005f_listtitle"] ||
            folderProperties["vti_listtitle"] ||
            folderProperties["listtitle"] ||
            folderProperties["title"];
          if (typeof listTitle !== "string")
            throw new Error("The dropped SharePoint folder did not provide a list title.");
          const listItemId = Number(itemKey[3]);
          if (!Number.isSafeInteger(listItemId))
            throw new Error("The dropped SharePoint file reference has an invalid list item ID.");

          const listItemResponse: unknown = await web.lists
            .getByTitle(listTitle)
            .items.getById(listItemId)
            .select("File")
            .expand("File")();
          if (!isSharePointFileListItem(listItemResponse))
            throw new TypeError("The dropped SharePoint file metadata has an invalid format.");
          const listItem = listItemResponse;
          const description = `Created: ${new Date(listItem.File.TimeCreated)} - Modified: ${new Date(listItem.File.TimeLastModified)}`;

          const fileBuffer: ArrayBuffer = await web
            .getFileByServerRelativePath(listItem.File.ServerRelativeUrl)
            .getBuffer();
          await this.appendFile(listItem.File.Name, fileBuffer, this.attachmentResults(), description);
        } catch (error: unknown) {
          errors.push(`SharePoint file append error: ${itemKey?.[2] ?? "unknown file"} with error ${MapErrors(error)}`);
        }
      }
    }

    if (errors.length > 0) {
      const errorMessages = errors.join("\n");
      alert(`Error retrieving item(s):\n\n${errorMessages}`);
      throw errorMessages;
    }
  }

  private officeAddin(): void {
    // dont double load the script
    if (document.getElementById("officejs") || this.officeApi) return;
    // try and determine if we are in an office addin
    try {
      if (
        "IsOfficeURLSchemes" in window ||
        window.location.search.includes("_host_Info=") ||
        window.location.search.includes("spc_addin=1")
      ) {
        throw "In an addin";
      }
      if (window.top != window.self) {
        throw "In an iframe, likely an addin";
      }
      // unlikely to be an addin
      return;
    } catch {
      // continue as its probably an addin
    }
    // load the office.js script as web/pwa apps are iframes and full clients are not with no obvious way to detect if its needed
    const officeScript = document.createElement("script");
    officeScript.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
    officeScript.id = "officejs";
    document.head.appendChild(officeScript);
    // capture the office type for later use along with triggering change detection if needed
    officeScript.addEventListener(
      "load",
      () => {
        // only once office.js loaded
        const office = this.officeApi;
        if (office?.onReady) {
          // on ready should be ran soon after office.js is loaded trigger
          office.onReady((info) => {
            // diagnostics for office addin
            console.log("Office", office);
            console.log("Office.onReady info", info);
            console.log("Office.nested", office.context?.requirements?.isSetSupported?.("NestedAppAuth", "1.1"));
            // capture what office type of addin for later use
            if (this.office.type != info.host.toString()) {
              this.office.type = info.host.toString();
              // Office.js may replace native dialogs with throwing stubs in add-ins.
              // Override them so callers continue; confirm approves because no native prompt is available.
              window.alert = (message?: unknown): void => {
                this.office.alert = message == null ? "" : String(message);
                this.chRef.markForCheck();
                this.chRef.detectChanges();
              };
              window.confirm = (message?: unknown): boolean => {
                this.office.alert = "Auto approving... " + (message == null ? "" : String(message));
                this.chRef.markForCheck();
                this.chRef.detectChanges();
                return true;
              };
              // abuse changed to mark for change detection and ensure running
              this.changed();
            }
          });
        }
      },
      false,
    );
  }

  // import from office addin selection or document panel
  private async importOutlook(): Promise<void> {
    const office = this.officeApi;
    if (!office) return;
    this.office.loading = true;
    this.office.alert = undefined;
    this.chRef.markForCheck();

    try {
      let targetMailbox: string | null = null;
      const mailbox = office.context.mailbox;
      const initialData =
        isRecord(mailbox) && isRecord(mailbox["initialData"]) ? mailbox["initialData"] : undefined;
      if (initialData?.["isFromSharedFolder"]) {
        targetMailbox = await new Promise<string>((resolve, reject) => {
          const item = office.context.mailbox.item;
          if (!item) {
            reject(new Error("The shared mailbox item is unavailable"));
            return;
          }
          item.getSharedPropertiesAsync((sharedProperties) => {
            if (sharedProperties.status === Office.AsyncResultStatus.Failed)
              reject(new Error(sharedProperties.error.message));
            else resolve(sharedProperties.value.targetMailbox);
          });
        });
      }

      await this.getMailItem(office, targetMailbox, new SharepointChoiceUtils());
    } catch (error) {
      this.office.alert = MapErrors(error, "Error retrieving mail(s):");
      this.errorHandler.handleError(error);
    } finally {
      this.office.loading = false;
      this.chRef.markForCheck();
    }
  }

  getMailItem = async (
    office: typeof Office,
    targetMailbox: string | null,
    sharepointUtils: SharepointChoiceUtils,
  ): Promise<void> => {
    const messages = await new Promise<Office.SelectedItemDetails[]>((resolve, reject) => {
      office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) reject(new Error(asyncResult.error.message));
        else resolve(asyncResult.value);
      });
    });

    const errors: string[] = [];
    for (const message of messages) {
      try {
        const fileContent = await sharepointUtils.callApi(
          undefined,
          undefined,
          undefined,
          targetMailbox == null
            ? `https://graph.microsoft.com/v1.0/me/messages/${message.itemId.replace(/\//g, "%252F")}/$value`
            : `https://graph.microsoft.com/v1.0/users/${targetMailbox}/messages/${message.itemId.replace(/\//g, "%252F")}/$value`,
          "GET",
          undefined,
          "text",
        );
        if (fileContent === null) throw new Error("The email request returned no message content.");

        await this.appendFile(
          `${message.subject.trim()}.eml`,
          new TextEncoder().encode(fileContent).buffer as ArrayBuffer,
          this.attachmentResults(),
        );
      } catch (error) {
        errors.push(`Email append error: ${message.subject} with error ${MapErrors(error)}`);
      }
    }

    if (errors.length) throw new Error(errors.join("\n"));
  };

  // import from office addin selection or document panel
  private importOffice() {
    const office = this.officeApi;
    if (!office) return;
    this.office.loading = true;
    this.office.alert = undefined;
    this.chRef.markForCheck();

    // if the adding type is word or excel then get the current document
    const docDataSlices: number[][] = [];
    let slicesReceived = 0;
    let docDataBytes = 0;
    let officeFile: Office.File | undefined;
    const closeFile = () => {
      const currentFile = officeFile;
      officeFile = undefined;
      currentFile?.closeAsync();
    };
    const failImport = (error: unknown) => {
      this.office.loading = false;
      closeFile();
      this.office.alert = MapErrors(error, "Error retrieving file(s):");
      this.errorHandler.handleError(error);
      this.chRef.markForCheck();
    };

    try {
      // get the file in 64k slices until complete
      office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 65536 },
        (asyncResult: Office.AsyncResult<Office.File>) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed)
            return failImport(new Error(asyncResult.error.message));
          try {
            officeFile = asyncResult.value;
            officeFile.getSliceAsync(0, processSlice);
          } catch (error) {
            failImport(error);
          }
        },
      );

      // process the slice and append until completed
      const processSlice = (sliceResult: Office.AsyncResult<Office.Slice>) => {
        if (sliceResult.status === Office.AsyncResultStatus.Failed)
          return failImport(new Error(sliceResult.error.message));

        try {
          const sliceData = sliceResult.value.data as number[];
          docDataBytes += sliceData.length;
          if (docDataBytes > SharepointChoiceComponent.maxFileSizeBytes)
            return failImport(new Error(`The open file exceeds the ${this.maxFileSizeLabel} file size limit`));

          docDataSlices[sliceResult.value.index] = sliceData;
          const currentFile = officeFile;
          if (!currentFile) return;

          if (++slicesReceived == currentFile.sliceCount) {
            void saveFile()
              .then(() => {
                closeFile();
                this.office.loading = false;
                this.chRef.markForCheck();
              })
              .catch(failImport);
          } else currentFile.getSliceAsync(slicesReceived, processSlice);
        } catch (error) {
          failImport(error);
        }
      };

      const saveFile = async () => {
        const documentData = new Uint8Array(docDataBytes);
        let offset = 0;
        for (const sliceData of docDataSlices) {
          documentData.set(sliceData ?? [], offset);
          offset += sliceData?.length ?? 0;
        }

        // try to get a file name from the document
        let fileName = office.context.document.url?.split("/")?.pop()?.split("\\")?.pop();
        if (!fileName) {
          fileName = `OfficeAddin-${new Date().toISOString().replace(/:/g, "-")}`;
          switch (this.office.type) {
            case "Word":
              fileName += ".docx";
              break;
            case "Excel":
              fileName += ".xlsx";
              break;
            case "PowerPoint":
              fileName += ".pptx";
              break;
            case "OneNote":
              fileName += ".one";
              break;
            case "Project":
              fileName += ".mpp";
              break;
            case "Visio":
              fileName += ".vsdx";
              break;
            default:
              throw "Unknown Office type, unable to save file";
          }
        }

        await this.appendFile(fileName, documentData.buffer, this.attachmentResults());
      };
    } catch (error) {
      failImport(error);
    }
  }

  private async appendFile(
    fileName: string | null | undefined,
    data: ArrayBuffer,
    results: SharepointChoiceAttachment[],
    description?: string,
    depth: number = 0,
  ) {
    const resolvedFileName = fileName?.trim() || `Attachment-${globalThis.crypto.randomUUID()}.bin`;

    if (data.byteLength > SharepointChoiceComponent.maxFileSizeBytes)
      throw new Error(`${resolvedFileName} exceeds the ${this.maxFileSizeLabel} file size limit`);

    // cleanup the name
    const sanitizedFileName = resolvedFileName
      .trim()
      .replace(/[+\\/:*?"%'#<>|={}~;@]/g, "-")
      .replace(/-+/g, "-")
      .replace(/\.+/g, ".")
      .replace(/[. -]+$/g, "")
      .replace(/^[. -]+/g, "")
      .replace(/[\u200B-\u200F\uFEFF]/g, "");
    const normalizedFileName = sanitizedFileName || `Attachment-${globalThis.crypto.randomUUID()}.bin`;
    const titleSource = sanitizedFileName ? resolvedFileName : normalizedFileName;

    // only save files that have a name and extension
    if (normalizedFileName.indexOf(".") < 0 || normalizedFileName.length < 3) return;
    // get the extension
    const fileExtension = normalizedFileName.substring(normalizedFileName.lastIndexOf(".") + 1);
    // get the first part of the name
    const baseFileName = normalizedFileName.substring(0, normalizedFileName.lastIndexOf(".")).substring(0, 100);
    // get the title
    const title = titleSource.trim().substring(0, titleSource.lastIndexOf(".")).substring(0, 255);
    // find the next available file name by appending a number
    let duplicateIndex = 0;
    let uniqueFileName = `${baseFileName}.${fileExtension}`;
    while (results.some((existingAttachment) => existingAttachment.FileName == uniqueFileName)) {
      uniqueFileName = `${baseFileName} (${++duplicateIndex}).${fileExtension}`;
    }

    const newAttachment: SharepointChoiceAttachment = {
      FileName: uniqueFileName,
      Data: data,
      Length: data.byteLength,
      ListItemAllFields: { Title: title },
    };

    if (description && this.file?.notes && newAttachment.ListItemAllFields)
      newAttachment.ListItemAllFields[this.file.notes] = description;

    results.push(newAttachment);

    const extension = fileExtension.toLowerCase();
    if (this.file?.extract && ["zip", "msg", "eml"].includes(extension)) {
      if (depth >= SharepointChoiceComponent.maxExtractionDepth) {
        this.errorHandler.handleError({
          level: "Warning",
          message: `File extraction issue for ${uniqueFileName}: Maximum extraction depth of ${SharepointChoiceComponent.maxExtractionDepth} reached`,
          undefined,
        });
      } else if (extension == "zip" && (await this.zips(data, results, uniqueFileName, depth))) {
        await this.delete(results.find((attachment) => attachment.FileName == uniqueFileName));
      } else if (extension == "msg") await this.msgs(data, results, uniqueFileName, depth);
      else if (extension == "eml") await this.emls(data, results, uniqueFileName, depth);
    }

    this.changed();
  }

  // extract zip files and append to results
  private async zips(data: ArrayBuffer, results: SharepointChoiceAttachment[], sourceFile: string, depth: number = 0) {
    const errors: string[] = [];
    let extractionFailures = 0;
    try {
      const archive = await loadAsync(data);
      const archiveFiles = Object.values(archive.files);
      const normalizeArchivePath = (path: string) =>
        path
          .replace(/(\.+[\\/])+/g, "")
          .replace(/^\.+/, "")
          .replace(/^[\\/]+/, "")
          .replace(/\.+$/, "")
          .replace(/[\\/]+$/, "")
          .replace(/[\\/]+/g, "/");
      // get common root folder to suppress some depth
      const pathSegments = archiveFiles
        .filter((archiveEntry) => !archiveEntry.dir)
        .map((archiveEntry) => normalizeArchivePath(archiveEntry.name).split("/").slice(0, -1));
      const commonRootSegments = pathSegments.length
        ? pathSegments.reduce((sharedSegments, candidateSegments) =>
            sharedSegments.filter((segment, index) => segment == candidateSegments[index]),
          )
        : [];
      const commonRootPathLength = (commonRootSegments.length > 0 ? commonRootSegments.join("/") + "/" : "").length;
      for (const archiveEntry of archiveFiles) {
        if (archiveEntry.dir) continue;

        try {
          const internalData: unknown = Reflect.get(archiveEntry, "_data");
          const uncompressedSize: unknown =
            internalData && typeof internalData === "object" && "uncompressedSize" in internalData
              ? Reflect.get(internalData, "uncompressedSize")
              : undefined;
          if (typeof uncompressedSize === "number" && uncompressedSize > SharepointChoiceComponent.maxFileSizeBytes)
            throw new Error(`${archiveEntry.name} exceeds the ${this.maxFileSizeLabel} file size limit`);

          const fileBuffer = await archiveEntry.async("arraybuffer");
          const flattenedName = normalizeArchivePath(archiveEntry.name).substring(commonRootPathLength);
          await this.appendFile(flattenedName, fileBuffer, results, `Date: ${archiveEntry.date}`, depth + 1);
        } catch (error: unknown) {
          extractionFailures++;
          errors.push(`Error extracting ZIP file ${archiveEntry.name} from ${sourceFile}: ${MapErrors(error)}`);
        }
      }
    } catch (error: unknown) {
      // zip is uploaded so any extracted elements are only nice to have
      extractionFailures++;
      errors.push(`Error extracting ZIP from ${sourceFile}: ${MapErrors(error)}`);
    }

    if (errors.length > 0) {
      this.errorHandler.handleError({
        level: "Warning",
        message: `File extraction issue for ${sourceFile}: ${errors.join("\n")}`,
        undefined,
      });
    }

    return extractionFailures == 0;
  }

  // extract and append msg email attachments to results
  private async msgs(data: ArrayBuffer, results: SharepointChoiceAttachment[], sourceFile: string, depth: number = 0) {
    const errors: string[] = [];

    try {
      // The package's CommonJS entry exposes its constructor on the default property at runtime.
      const msgReaderConstructor = (MsgReader as typeof MsgReader & { default: typeof MsgReader })["default"];
      const msgReader = new msgReaderConstructor(data);
      // triggered the parser
      const messageData = msgReader.getFileData();
      // if no sender name then its not an email
      if (!("senderName" in messageData)) throw new Error("Not a valid MSG email file");
      // get the email date
      const dateHeaders = messageData.headers?.split("\n").filter((header: string) => header.startsWith("Date: "));
      const dateHeader = dateHeaders?.[0];
      const receivedDate = dateHeader ? new Date(dateHeader.replace("Date: ", "")) : new Date();
      // get all attachments
      for (const attachment of messageData.attachments || []) {
        if (attachment.attachmentHidden) continue;

        try {
          const attachmentFile = msgReader.getAttachment(attachment);
          await this.appendFile(
            attachmentFile.fileName,
            attachmentContentBuffer(attachmentFile.content),
            results,
            `Sent: ${receivedDate}`,
            depth + 1,
          );
        } catch (error: unknown) {
          errors.push(`Error extracting MSG attachment from ${sourceFile}: ${MapErrors(error)}`);
        }
      }
    } catch (error: unknown) {
      errors.push(`Error extracting MSG from ${sourceFile}: ${MapErrors(error)}`);
    }

    if (errors.length > 0) {
      this.errorHandler.handleError({
        level: "Warning",
        message: `File extraction issue for ${sourceFile}: ${errors.join("\n")}`,
        undefined,
      });
    }
  }

  // extract and append eml email attachments to results
  private async emls(data: ArrayBuffer, results: SharepointChoiceAttachment[], sourceFile: string, depth: number = 0) {
    const errors: string[] = [];

    try {
      const email = await PostalMime.parse(data, { attachmentEncoding: "arraybuffer" });

      if (!email || !email.from) throw new Error("Not a valid EML email file");

      const receivedDate = email.date ? new Date(email.date) : new Date();

      for (const attachment of email.attachments as PostalMimeAttachment[]) {
        if (attachment.disposition === "inline" || attachment.related) continue;

        try {
          // work out the name from content id which is more consistent across sources, otherwise from filename.
          const contentIdFileName = attachment.contentId?.replace(/^</, "").replace(/>$/, "").split("@")[0];
          const attachmentFileName =
            contentIdFileName && contentIdFileName.indexOf(".") >= 0 ? contentIdFileName : attachment.filename;

          await this.appendFile(
            attachmentFileName,
            attachmentContentBuffer(attachment.content),
            results,
            `Sent: ${receivedDate}`,
            depth + 1,
          );
        } catch (error: unknown) {
          errors.push(`Error extracting EML attachment from ${sourceFile}: ${MapErrors(error)}`);
        }
      }
    } catch (error: unknown) {
      errors.push(`Error extracting EML from ${sourceFile}: ${MapErrors(error)}`);
    }

    if (errors.length > 0) {
      this.errorHandler.handleError({
        level: "Warning",
        message: `File extraction issue for ${sourceFile}: ${errors.join("\n")}`,
        undefined,
      });
    }
  }

  // dragging and dropping, hover
  private over(event: DragEvent) {
    if (this.disabled) return;
    event.preventDefault();
    event.stopPropagation();
    this.filesOver = true;
    this.chRef.markForCheck();
  }

  // dragging and dropping, unhover
  private leave(event: DragEvent) {
    if (this.disabled) return;
    event.preventDefault();
    event.stopPropagation();
    this.filesOver = false;
    this.chRef.markForCheck();
  }

  // dragging and dropping, drop
  private async drop(event: DragEvent) {
    if (!event) return;
    if (this.disabled) return;
    event.preventDefault();
    event.stopPropagation();

    try {
      if (event.dataTransfer?.files?.length) await this.add(event.dataTransfer);
      if (event.dataTransfer?.items?.length) await this.outlook(event.dataTransfer);

      // if no transfer on drop and not chromium based add the meta tag to allow the drop next time
      if (!event.dataTransfer && !(window as SharepointChoiceWindow).chrome) {
        const compatibilityMeta = document.createElement("meta");
        compatibilityMeta.httpEquiv = "X-UA-Compatible";
        compatibilityMeta.content = "chrome=1";
        document.head.appendChild(compatibilityMeta);
      }
    } catch (error) {
      this.errorHandler.handleError(error);
    } finally {
      this.filesOver = false;
      this.chRef.markForCheck();
    }
  }

  /*
  Common parts between user field types
  */

  // select user
  private async selectedUser(selectedUser: SharepointChoiceUser | null): Promise<void> {
    this.userSearchRequestId++;
    if (
      selectedUser &&
      !this.displayedUsers.some((displayedUser) => displayedUser.LoginName == selectedUser.LoginName)
    ) {
      const sharepointUtils = new SharepointChoiceUtils();
      const userScopeUrl = this.spec?.[this.field]?.Scope || "";
      const web = Web([sharepointUtils.sp.web, userScopeUrl]);
      const ensureResponse: unknown = await web.ensureUser(selectedUser.LoginName);
      if (
        !isRecord(ensureResponse) ||
        typeof ensureResponse["Id"] !== "number" ||
        !Number.isSafeInteger(ensureResponse["Id"])
      )
        throw new TypeError("SharePoint did not return a valid ID for the selected user.");
      const ensuredUser = {
        Id: ensureResponse["Id"],
        Title: selectedUser.Title,
        LoginName: selectedUser.LoginName,
      };

      // add to field
      if (this.get("TypeAsString") == "UserMulti") {
        const results = readResults(this.form[this.field + "Id"]) ?? [];
        results.push(ensuredUser.Id);
        setResults(this.form, this.field + "Id", results);
        this.displayedUsers.push(ensuredUser);
      } else {
        this.form[this.field + "Id"] = ensuredUser.Id;
        this.displayedUsers = [ensuredUser];
      }
    }

    // clear search fields
    this.otherText = "";
    this.users = [];

    this.changed();
  }

  // load list data only has IDs so expand the object
  private displayUser(user?: number): string {
    if (!user || user <= 0) return "";

    const displayedUser = this.displayedUsers.find((candidate) => candidate.Id == user);
    if (displayedUser?.Title) return displayedUser.Title;

    const retryAt = this.userLoadRetryAt.get(user);
    if (retryAt && Date.now() < retryAt) return "";
    this.userLoadRetryAt.delete(user);

    // dont trigger a new load web request if the users aready loading
    if (!this.loadingUserIds.includes(user)) {
      this.loadingUserIds.push(user);
      // load the user
      const sharepointUtils = new SharepointChoiceUtils();
      const userScopeUrl = this.spec?.[this.field]?.Scope || "";
      const web = Web([sharepointUtils.sp.web, userScopeUrl]);
      // do as .then to avoid async issues
      web
        .getUserById(user)()
        .then((response: unknown) => {
          if (this.destroyed) return;
          if (!isSharepointChoiceUser(response)) throw new TypeError("SharePoint returned invalid user details.");
          const loadedUser = response;

          // update the display table
          this.displayedUsers.push({
            LoginName: loadedUser.LoginName,
            Id: loadedUser.Id,
            Title: loadedUser.Title,
          });

          this.chRef.markForCheck();
        })
        .catch((error: unknown) => {
          if (this.destroyed) return;
          this.userLoadRetryAt.set(user, Date.now() + 5000);
          this.errorHandler.handleError(error);
        })
        .finally(() => {
          const loadingIndex = this.loadingUserIds.indexOf(user);
          if (loadingIndex >= 0) this.loadingUserIds.splice(loadingIndex, 1);
          if (!this.destroyed) this.chRef.markForCheck();
        });
    }

    return "";
  }

  // removes a user
  private removeUser(userId: number | null): void {
    if (!userId) {
      this.form[this.field + "Id"] = null;
      this.displayedUsers = [];
    } else {
      const results = readResults(this.form[this.field + "Id"]);
      if (results)
        setResults(
          this.form,
          this.field + "Id",
          results.filter((value) => value !== userId),
        );
      this.displayedUsers = this.displayedUsers.filter((displayedUser) => displayedUser.Id != userId);
    }
    this.changed();
  }

  // trigger user search
  private onUpUser(pressedKey: string | undefined): void {
    if (pressedKey == "ArrowDown") {
      if (this.pos < this.users.length - 1) this.pos++;
    } else if (pressedKey == "ArrowUp") {
      if (this.pos > 0) this.pos--;
    } else if (pressedKey == "Enter") {
      this.userSearchRequestId++;
      this.selectedUser(this.users[this.pos] ?? null);
    } else {
      if (pressedKey !== undefined) {
        this.userSearchRequestId++;
        this.users = [];
        this.pos = -1;
      }
      if (!this.otherText || this.otherText.length < 3) this.users = [];
      else this.userKey?.next(this.otherText);
    }
    this.chRef.markForCheck();
  }

  private async onUpUserSearch(searchText: string): Promise<void> {
    if (!searchText) return;
    const requestId = ++this.userSearchRequestId;

    const searchQuery = {
      // set the user partial being searched
      QueryString: searchText,
      MaximumEntitySuggestions: 10,
      AllowEmailAddresses: true,
      AllowOnlyEmailAddresses: false,
      PrincipalSource: 15,
      PrincipalType: 1,
      // set group each query to adapt to changes
      SharePointGroupID: parseInt(String(this.get("SelectionGroup") ?? 0), 10),
    };

    try {
      const sharepointUtils = new SharepointChoiceUtils();
      const response: unknown = await sharepointUtils.sp.profiles.clientPeoplePickerSearchUser(searchQuery);
      const matchingUsers = normalizePeoplePickerResults(response);
      if (requestId !== this.userSearchRequestId) return;

      this.users = [];
      this.pos = -1;
      this.users = matchingUsers;
    } catch (error) {
      if (requestId !== this.userSearchRequestId) return;
      this.users = [];
      this.pos = -1;
      this.errorHandler.handleError(error);
    } finally {
      if (requestId === this.userSearchRequestId) this.chRef.markForCheck();
    }
  }
}
