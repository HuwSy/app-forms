import {
  Component,
  Input,
  ErrorHandler,
  EventEmitter,
  Output,
  ChangeDetectorRef,
  ChangeDetectionStrategy,
  OnInit,
  OnDestroy,
  isDevMode,
} from "@angular/core";
import { CommonModule } from "@angular/common";
import { FormsModule } from "@angular/forms";
import {
  SharepointChoiceColumn,
  SharepointChoiceTabs,
  SharepointChoiceRow,
  SharepointChoiceForm,
  SharepointChoiceField,
  SharepointChoiceList,
  SharepointChoiceRowChild,
  SharepointChoiceExportColumn,
  SharepointChoiceExportContext,
  SharepointChoiceExportOptions,
  ExcelIcon,
  SharepointChoiceRowValue,
  SharepointChoiceRowClickHandler,
  SharepointChoiceHyperlinkRowHandler,
  SharepointChoiceChangeEvent,
} from "./sharepoint-choice.models";
import { SharepointChoiceComponent } from "./sharepoint-choice.component";
import { SharepointChoiceLogging } from "./sharepoint-choice.logging";
import { SharepointChoiceRender, SharepointChoiceRenderContent } from "./sharepoint-choice.render";
import { Workbook } from "devextreme-exceljs-fork";

interface SharepointChoiceHide {
  [tabName: string]: string[];
}

interface SharepointChoiceSort {
  [tabName: string]: {
    direction: "asc" | "desc";
    field: string;
  }[];
}

interface SharepointChoiceFilter {
  [tabName: string]: {
    [fieldName: string]: {
      equals?: Array<string> | string | number | boolean | Date | null;
      contains?: string | null;
      greater?: number | Date | null;
      less?: number | Date | null;
    };
  };
}

type SharepointChoiceFilterOperation = "equals" | "contains" | "greater" | "less";

interface SharepointChoiceOrder {
  [tabName: string]: string[];
}

type SharepointChoiceDistinctValue = Exclude<SharepointChoiceRowValue, null | undefined>;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isSharepointChoiceColumn(value: unknown): value is SharepointChoiceColumn {
  if (!isRecord(value)) return false;

  if (
    (value["field"] !== undefined && typeof value["field"] !== "string") ||
    (value["headerName"] !== undefined && typeof value["headerName"] !== "string") ||
    (value["headerTooltip"] !== undefined && typeof value["headerTooltip"] !== "string") ||
    (value["width"] !== undefined && (typeof value["width"] !== "number" || !Number.isFinite(value["width"]))) ||
    (value["nowrap"] !== undefined && typeof value["nowrap"] !== "boolean") ||
    (value["center"] !== undefined && typeof value["center"] !== "boolean") ||
    (value["sortable"] !== undefined && typeof value["sortable"] !== "boolean") ||
    (value["_filtervisible"] !== undefined && typeof value["_filtervisible"] !== "boolean") ||
    (value["hide"] !== undefined && typeof value["hide"] !== "boolean" && typeof value["hide"] !== "function") ||
    (value["cellClicked"] !== undefined && typeof value["cellClicked"] !== "function") ||
    (value["cellRenderer"] !== undefined && typeof value["cellRenderer"] !== "function") ||
    (value["spec"] !== undefined && !isRecord(value["spec"]))
  ) {
    return false;
  }

  const filter = value["filter"];
  if (
    filter !== undefined &&
    filter !== "text" &&
    filter !== "number" &&
    filter !== "date" &&
    filter !== "select" &&
    filter !== "split" &&
    filter !== "none"
  ) {
    return false;
  }

  const children = value["children"];
  return children === undefined || (Array.isArray(children) && children.every(isSharepointChoiceColumn));
}

@Component({
  selector: "app-table",
  templateUrl: "./sharepoint-choice.table.html",
  styleUrls: ["./styles.scss"],
  standalone: true,
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [CommonModule, FormsModule, SharepointChoiceComponent, SharepointChoiceRender],
  providers: [
    {
      provide: ErrorHandler,
      useClass: SharepointChoiceLogging,
    },
  ],
})
export class SharepointChoiceTable implements OnInit, OnDestroy {
  // all data passed in keyed by tab name (not via signals which reduce performance on large data sets)
  @Input() set allData(value: SharepointChoiceTabs) {
    this._allDataIn = this.normalizeAllData(value);
    this._allData = {};
    this._distinctCache.clear();

    this._dataLoadCycles++;
    this._refreshCycle++;
    for (const tab of Object.keys(this._allDataIn)) {
      if (this._allDataIn[tab]) {
        // add _tracking to each row for ngFor tracking and node cache
        this._allDataIn[tab].forEach((row, index) => {
          if (row["_tracking"]) return;

          if (row["Id"])
            // likely SPId so retain through changes
            row["_tracking"] = `id-${row["Id"]}`;
          else if (row["reference"])
            // likely policy reference so retain through changes
            row["_tracking"] = `ref-${row["reference"]}`; // generate unique tracking key each time data changes too
          else row["_tracking"] = `${tab}-${this._dataLoadCycles}-${index}`;
        });
        // cleanup all data to only external filtered
        const filtered = !this.search
          ? this._allDataIn[tab]
          : this._allDataIn[tab].filter((row) => this.matches(row, this.search!));
        // only have tabs for ones with data
        if (this.showEmptyTabs || filtered.length > 0) this._allData[tab] = filtered;
      }
    }

    // Clear cache when data changes
    this.invalidateRows(true);
    this.chRef.markForCheck();
  }
  get allData(): SharepointChoiceTabs {
    return this._allData;
  }
  private _allDataIn: SharepointChoiceTabs = {};
  private _allData: SharepointChoiceTabs = {};

  // all columns
  @Input() set allCols(value: SharepointChoiceColumn[]) {
    this._allCols = this.normalizeColumns(value);
    // Clear cache when columns change
    this._colsCache.clear();
    this._nodeCache.clear();
    this.chRef.markForCheck();
  }
  get allCols(): SharepointChoiceColumn[] {
    return this._allCols;
  }
  private _allCols: SharepointChoiceColumn[] = [];

  // all tabs or derive from all data
  @Input() set allTabs(value: string[]) {
    const previousTab = this.selectedTab;
    this._allTabs = this.normalizeTabs(value);
    if (this.selectedTab !== previousTab) {
      this.pageNumber = 1;
      this._nodeCache.clear();
    }
    this.chRef.markForCheck();
  }
  get allTabs(): string[] {
    return this._allTabs || Object.keys(this.allData) || [];
  }
  private _allTabs?: string[];

  // selected tab else use stored value or first tab
  @Input() set selectedTab(value: string | undefined) {
    if (value != null && typeof value !== "string") {
      this.reportInvalidInput("selectedTab", "a string or undefined", value);
      value = undefined;
    }

    const previousTab = this.selectedTab;
    this._selectedTab = value;
    this.setStorage(`Tab`, this._selectedTab);
    // Revert to page 1 on every tab change
    this.pageNumber = 1;
    if (this.selectedTab !== previousTab) this._nodeCache.clear();
    this.chRef.markForCheck();
  }
  get selectedTab(): string | undefined {
    if (this._selectedTab && this.allTabs.includes(this._selectedTab)) return this._selectedTab;
    else if (this.allTabs.length > 0) return this.allTabs[0];
    return undefined;
  }
  private _selectedTab?: string;

  // page size via getter/setter to allow storage
  @Input() set pageSize(value: number | string) {
    this._pageSize = this.normalizePageSize(value);
    this.setStorage(`Size`, this._pageSize);
    this.pageNumber = 1;
    this.chRef.markForCheck();
  }
  get pageSize(): number {
    return this._pageSize ?? 100;
  }
  private _pageSize?: number;

  private normalizePageSize(value: unknown): number {
    const parsed =
      typeof value === "number" ? value : typeof value === "string" && value.trim() ? Number(value) : Number.NaN;
    if (Number.isFinite(parsed) && parsed >= 1) return Math.floor(parsed);
    if (value !== undefined && value !== null && value !== "") {
      this.reportInvalidInput("pageSize", "a finite number greater than or equal to 1", value);
    }
    return 100;
  }

  private reportInvalidInput(path: string, expected: string, value: unknown): void {
    if (!isDevMode()) return;
    const actual = value === null ? "null" : Array.isArray(value) ? "array" : typeof value;
    console.error(`[SharepointChoiceTable] Invalid ${path}: expected ${expected}; received ${actual}.`);
  }

  private reportRuntimeError(context: string, error: unknown): void {
    if (isDevMode()) console.error(`[SharepointChoiceTable] ${context} failed.`, error);
  }

  private normalizeAllData(value: unknown): SharepointChoiceTabs {
    if (!isRecord(value)) {
      this.reportInvalidInput("allData", "an object of tab names to row arrays", value);
      return {};
    }

    const result: SharepointChoiceTabs = {};
    for (const [tab, rows] of Object.entries(value)) {
      if (!Array.isArray(rows)) {
        this.reportInvalidInput(`allData["${tab}"]`, "an array of row objects", rows);
        continue;
      }

      const validRows: SharepointChoiceRow[] = [];
      rows.forEach((row, index) => {
        if (!isRecord(row)) {
          this.reportInvalidInput(`allData["${tab}"][${index}]`, "a row object", row);
          return;
        }
        // List columns can have schema-specific nested values; preserve them for configured columns to interpret.
        validRows.push(row as SharepointChoiceRow);
      });
      result[tab] = validRows;
    }
    return result;
  }

  private normalizeColumns(value: unknown): SharepointChoiceColumn[] {
    if (!Array.isArray(value)) {
      this.reportInvalidInput("allCols", "an array of column definitions", value);
      return [];
    }

    return value.filter((column, index): column is SharepointChoiceColumn => {
      if (isSharepointChoiceColumn(column)) return true;
      this.reportInvalidInput(`allCols[${index}]`, "a valid column definition", column);
      return false;
    });
  }

  private normalizeTabs(value: unknown): string[] | undefined {
    if (value == null) return undefined;
    if (!Array.isArray(value)) {
      this.reportInvalidInput("allTabs", "an array of strings or undefined", value);
      return undefined;
    }

    return value.filter((tab, index): tab is string => {
      if (typeof tab === "string") return true;
      this.reportInvalidInput(`allTabs[${index}]`, "a string", tab);
      return false;
    });
  }

  private normalizeStringArrayMap(value: unknown, setting: string): Record<string, string[]> {
    if (value == null) return {};
    if (!isRecord(value)) {
      this.reportInvalidInput(setting, "an object of string arrays", value);
      return {};
    }

    const result: Record<string, string[]> = {};
    for (const [tab, values] of Object.entries(value)) {
      if (!Array.isArray(values)) {
        this.reportInvalidInput(`${setting}["${tab}"]`, "an array of strings", values);
        continue;
      }
      result[tab] = values.filter((item, index): item is string => {
        if (typeof item === "string") return true;
        this.reportInvalidInput(`${setting}["${tab}"][${index}]`, "a string", item);
        return false;
      });
    }
    return result;
  }

  private normalizeSort(value: unknown): SharepointChoiceSort {
    if (value == null) return {};
    if (!isRecord(value)) {
      this.reportInvalidInput("sort", "an object of sort arrays", value);
      return {};
    }

    const result: SharepointChoiceSort = {};
    for (const [tab, entries] of Object.entries(value)) {
      if (!Array.isArray(entries)) {
        this.reportInvalidInput(`sort["${tab}"]`, "an array of sort definitions", entries);
        continue;
      }

      result[tab] = entries.filter((entry, index): entry is SharepointChoiceSort[string][number] => {
        if (
          isRecord(entry) &&
          typeof entry["field"] === "string" &&
          (entry["direction"] === "asc" || entry["direction"] === "desc")
        ) {
          return true;
        }
        this.reportInvalidInput(`sort["${tab}"][${index}]`, "a field name and asc/desc direction", entry);
        return false;
      });
    }
    return result;
  }

  private normalizeFilter(value: unknown): SharepointChoiceFilter {
    if (value == null) return {};
    if (!isRecord(value)) {
      this.reportInvalidInput("filter", "an object of tab filters", value);
      return {};
    }

    const result: SharepointChoiceFilter = {};
    for (const [tab, fields] of Object.entries(value)) {
      if (!isRecord(fields)) {
        this.reportInvalidInput(`filter["${tab}"]`, "an object of field filters", fields);
        continue;
      }

      const validFields: SharepointChoiceFilter[string] = {};
      for (const [field, operations] of Object.entries(fields)) {
        if (!isRecord(operations)) {
          this.reportInvalidInput(`filter["${tab}"]["${field}"]`, "an object of filter operations", operations);
          continue;
        }

        const validOperations: SharepointChoiceFilter[string][string] = {};
        for (const [operation, filterValue] of Object.entries(operations)) {
          const path = `filter["${tab}"]["${field}"]["${operation}"]`;
          if (filterValue === undefined) continue;

          if (operation === "equals") {
            if (
              filterValue === null ||
              typeof filterValue === "string" ||
              typeof filterValue === "boolean" ||
              (typeof filterValue === "number" && Number.isFinite(filterValue)) ||
              (filterValue instanceof Date && Number.isFinite(filterValue.getTime()))
            ) {
              validOperations.equals = filterValue;
            } else if (Array.isArray(filterValue)) {
              const strings = filterValue.filter((item): item is string => typeof item === "string");
              if (strings.length === filterValue.length) validOperations.equals = strings;
              else this.reportInvalidInput(path, "a scalar or array of strings", filterValue);
            } else {
              this.reportInvalidInput(path, "a scalar or array of strings", filterValue);
            }
          } else if (operation === "contains") {
            if (filterValue === null || typeof filterValue === "string") validOperations.contains = filterValue;
            else this.reportInvalidInput(path, "a string or null", filterValue);
          } else if (operation === "greater" || operation === "less") {
            const rangeValue = this.normalizeRangeFilter(filterValue, path);
            if (rangeValue !== undefined) validOperations[operation] = rangeValue;
          } else {
            this.reportInvalidInput(path, "equals, contains, greater, or less", filterValue);
          }
        }

        if (Object.keys(validOperations).length > 0) validFields[field] = validOperations;
      }
      result[tab] = validFields;
    }
    return result;
  }

  private normalizeRangeFilter(value: unknown, path: string): number | Date | null | undefined {
    if (value === null) return null;
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (value instanceof Date && Number.isFinite(value.getTime())) return value;
    if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}(?:T.*)?$/.test(value)) {
      const date = new Date(value);
      if (Number.isFinite(date.getTime())) return date;
    }

    this.reportInvalidInput(path, "a finite number, valid date, or ISO date string", value);
    return undefined;
  }

  // should show loading state and disable clicks etc
  @Input() set loading(value: boolean) {
    if (typeof value !== "boolean" && value != null) {
      this.reportInvalidInput("loading", "a boolean", value);
    }
    this._loading = value === true;
    this.invalidateRows();
    this.chRef.markForCheck();
  }
  get loading(): boolean {
    return this._loading;
  }
  private _loading: boolean = false;

  @Input() set search(value: SharepointChoiceRowChild | undefined) {
    if (value != null && !isRecord(value)) {
      this.reportInvalidInput("search", "an object or undefined", value);
      value = undefined;
    }
    this._search = value;

    this.sort = this.normalizeSort(this.getStorage(`Sort`));
    this.filter = this.normalizeFilter(this.getStorage(`Filter`));
    this.columnOrder = this.normalizeStringArrayMap(this.getStorage(`Order`), "columnOrder");
    this.hiddenColumns = this.normalizeStringArrayMap(this.getStorage(`Hide`), "hiddenColumns");
    const storedTab: unknown = this.getStorage(`Tab`);
    if (storedTab == null || typeof storedTab === "string") this.selectedTab = storedTab ?? undefined;
    else {
      this.reportInvalidInput("selectedTab", "a string or undefined", storedTab);
      this.selectedTab = undefined;
    }
    const storedPageSize: unknown = this.getStorage(`Size`);
    if (typeof storedPageSize === "number" || typeof storedPageSize === "string") this.pageSize = storedPageSize;
    else if (storedPageSize != null)
      this.reportInvalidInput("pageSize", "a finite number greater than or equal to 1", storedPageSize);

    // trigger setter to apply search filter to all data and reset caches
    this.allData = this._allDataIn;
    this.invalidateRows(true);

    this.debounceAndMark();
  }
  get search(): SharepointChoiceRowChild | undefined {
    return this._search;
  }
  private _search?: SharepointChoiceRowChild;

  // Change this input after mutating row data in place so OnPush checks and caches are refreshed.
  @Input() set refreshKey(value: unknown) {
    this._refreshKey = value;
    this._refreshCycle++;
    this.invalidateRows(true);
    this._nodeCache.clear();
    this._distinctCache.clear();
    this.chRef.markForCheck();
  }
  get refreshKey(): unknown {
    return this._refreshKey;
  }
  private _refreshKey?: unknown;

  // simple inputs that dont need getter/setter
  @Input() prefix: string = document.location.href.toLowerCase().split("?")[0]!.split("#")[0]!;
  @Input() tableHeight: string = "calc(100vh - 310px)";
  @Input() allEditing: boolean = false; // all with spec render as app-choice else edit per cell on click
  @Input() allowHideColumns: boolean = true;
  @Input() showEmptyTabs: boolean = false;

  // allow selection and emit selected items and tab
  @Input() allowSelection: boolean = false;
  @Output() selected = new EventEmitter<{
    data: SharepointChoiceRow[];
    tab: string | undefined;
  }>();
  @Output() cleared = new EventEmitter<void>();

  // outbound events or pseudo callbacks for await support on click, in order of trigger
  // cellClicked triggers ahead of all of these if present
  // Keep the legacy callback inputs loosely typed for consumer compatibility; use explicit signatures for new APIs.
  @Input() rowClicked?: SharepointChoiceRowClickHandler; // [rowClicked]="rowClicked" rowClicked = async (rowData, target) => { ... } to ensure this. is from the app and not from app-table
  @Output() clicked = new EventEmitter<{
    row: SharepointChoiceRow;
    target: HTMLElement | EventTarget | undefined;
  }>(); // (clicked)="onClicked($event)" onClicked(event: { row: SharepointChoiceRow, target: HTMLElement | EventTarget | undefined }) { ... }
  @Input() hyperlinkRow?: SharepointChoiceHyperlinkRowHandler; // [hyperlinkRow]="hyperlinkRow" hyperlinkRow = (rowData) => { return 'https://...'; } to ensure this. is from the app and not from app-table
  @Input() hyperlinkTarget: string = "_self"; // target for hyperlink rows
  @Input() export?: SharepointChoiceExportOptions; // [export]="export" declarative export options handled by app-table

  // internal state
  set pageNumber(value: number) {
    this._pageNumber = value;
    this._pageCache = [];
    this.chRef.markForCheck();
  }
  get pageNumber(): number {
    return this._pageNumber || 1;
  }
  private _pageNumber?: number;

  editColumns = false;
  excelIcon: string = ExcelIcon;

  // stored sort/filter/hidden columns via getter/setter with storage and recall
  set sort(value: SharepointChoiceSort) {
    this._sort = this.normalizeSort(value);
    this.setStorage(`Sort`, this.sort);
    this.debounceAndMark();
  }
  get sort(): SharepointChoiceSort {
    return this._sort || {};
  }
  private _sort?: SharepointChoiceSort;

  set filter(value: SharepointChoiceFilter) {
    this._filter = this.normalizeFilter(value);
    this.setStorage(`Filter`, this.filter);
    this.pageNumber = 1;
    this.debounceAndMark();
  }
  get filter(): SharepointChoiceFilter {
    return this._filter || {};
  }
  private _filter?: SharepointChoiceFilter;

  set columnOrder(value: SharepointChoiceOrder) {
    this._columnOrder = this.normalizeStringArrayMap(value, "columnOrder");
    this.setStorage(`Order`, this.columnOrder);
    this._colsCache.clear();
    this.chRef.markForCheck();
  }
  get columnOrder(): SharepointChoiceOrder {
    return this._columnOrder || {};
  }
  private _columnOrder?: SharepointChoiceOrder;

  set hiddenColumns(value: SharepointChoiceHide) {
    this._hiddenColumns = this.normalizeStringArrayMap(value, "hiddenColumns");
    this.setStorage(`Hide`, this.hiddenColumns);
    // rebuild all on change as the ticks can be any
    this._colsCache.clear();
    this.invalidateRows(true);
    this.chRef.markForCheck();
  }
  get hiddenColumns(): SharepointChoiceHide {
    return this._hiddenColumns || {};
  }
  private _hiddenColumns?: SharepointChoiceHide;

  private _dataLoadCycles: number = 0;
  private _refreshCycle: number = 0;
  private _debounceFilterSort?: ReturnType<typeof setTimeout>;
  private _resizeCleanup?: () => void;

  // Memoization cache for columns filtered/sorted rows
  private _colsCache: Map<string, SharepointChoiceColumn[]> = new Map();
  private _rowsCache: Map<string, SharepointChoiceRow[]> = new Map();
  private _distinctCache: Map<string, SharepointChoiceDistinctValue[]> = new Map();
  private _pageCache: SharepointChoiceRow[] = [];
  private _fieldMapCache: Map<string, Array<string>> = new Map();
  private _nodeCache: Map<string, string | Node | null | undefined> = new Map();
  private _isObserved: boolean = false;
  private _dragColumn?: SharepointChoiceColumn;
  private _dragParent?: SharepointChoiceColumn;
  private _suppressHeaderClick: boolean = false;

  public constructor(
    private chRef: ChangeDetectorRef,
    private errorHandler: ErrorHandler,
  ) {}

  public ngOnInit(): void {
    if (!this._sort) this._sort = this.normalizeSort(this.getStorage(`Sort`));
    if (!this._filter) this._filter = this.normalizeFilter(this.getStorage(`Filter`));
    if (!this._columnOrder) this._columnOrder = this.normalizeStringArrayMap(this.getStorage(`Order`), "columnOrder");
    if (!this._hiddenColumns)
      this._hiddenColumns = this.normalizeStringArrayMap(this.getStorage(`Hide`), "hiddenColumns");
    if (!this._selectedTab) {
      const tab: unknown = this.getStorage(`Tab`);
      if (tab == null || typeof tab === "string") this._selectedTab = tab ?? undefined;
      else {
        this.reportInvalidInput("selectedTab", "a string or undefined", tab);
        this._selectedTab = undefined;
      }
    }
    if (!this._pageSize)
      try {
        const storedPageSize: unknown = this.getStorage(`Size`);
        if (typeof storedPageSize === "number" || typeof storedPageSize === "string") this.pageSize = storedPageSize;
        else if (storedPageSize != null)
          this.reportInvalidInput("pageSize", "a finite number greater than or equal to 1", storedPageSize);
      } catch (error) {
        console.log("SharepointChoiceTable.ngOnInit storage load error: Size", error);
      }
  }

  public ngOnDestroy(): void {
    clearTimeout(this._debounceFilterSort);
    this._resizeCleanup?.();
  }

  private matches(row: SharepointChoiceRow, search: SharepointChoiceRowChild): boolean {
    for (const key in search) {
      const searchValue = search[key];
      if (searchValue === null || searchValue === undefined) continue;

      if (
        searchValue instanceof Date === false &&
        typeof searchValue === "object" &&
        Array.isArray(searchValue) === false
      ) {
        const rowValue = row[key];
        if (!rowValue || typeof rowValue !== "object" || Array.isArray(rowValue)) return false;
        if (!this.matches(rowValue as SharepointChoiceRow, searchValue)) return false;
        continue;
      }

      const rowValue = row[key];
      if (searchValue === "(blanks)") {
        if (rowValue !== null && rowValue !== undefined && rowValue !== "") return false;
        continue;
      }

      if (rowValue === null || rowValue === undefined || rowValue === "") return false;
      if (!rowValue.toString().toLowerCase().includes(searchValue.toString().toLowerCase())) return false;
    }

    return true;
  }

  private invalidateRows(resetPage: boolean = false): void {
    this._rowsCache.clear();
    this._pageCache = [];
    if (resetPage && this.pageNumber !== 1) this.pageNumber = 1;
    this.chRef.markForCheck();
  }

  private debounceAndMark(resetPage: boolean = false): void {
    // use debounce to avoid multiple rapid changes clearing cache too often across multiple columns and filter/sort types etc so simpler than rxjs
    this.invalidateRows(resetPage);
    clearTimeout(this._debounceFilterSort);
    // Let nested controls and related filter/sort changes settle before emitting selection.
    this._debounceFilterSort = setTimeout(() => {
      if (!this.selectedTab) return;
      // drop cache to recalc rows
      this._rowsCache.delete(this.selectedTab);
      // emit the new filtered/sorted selection
      this.emitSelection();
    }, 500);
  }

  private storageKey(): string {
    return `SharepointTable-${this.prefix}`;
  }

  private storageObject(): Record<string, unknown> {
    const key = this.storageKey();

    try {
      const raw = localStorage.getItem(key);
      if (!raw) return {};

      try {
        const parsed: unknown = JSON.parse(raw);
        if (isRecord(parsed)) return parsed;
        this.reportInvalidInput(`localStorage["${key}"]`, "a JSON object", parsed);
      } catch (error) {
        this.reportRuntimeError(`Parsing table preferences "${key}"`, error);
      }
    } catch (error) {
      this.reportRuntimeError(`Reading table preferences "${key}"`, error);
      return {};
    }

    try {
      localStorage.setItem(key, "{}");
    } catch (error) {
      this.reportRuntimeError(`Resetting table preferences "${key}"`, error);
    }

    return {};
  }

  private getStorage(key: string): unknown {
    return this.storageObject()[key];
  }

  private setStorage(key: string, value: unknown): void {
    try {
      const state = this.storageObject();
      state[key] = value;
      localStorage.setItem(this.storageKey(), JSON.stringify(state));
    } catch (error) {
      this.reportRuntimeError(`Saving table preference "${key}"`, error);
    }
  }

  // tab change resets page number and emits empty selection
  private tabChange(tab: string): void {
    this.selectedTab = tab;
    this.emitSelection();
  }

  // selection change toggles row selected state and emits selected rows
  private selectionChanged(row: SharepointChoiceRow): void {
    row._selected = !row._selected;
    // if there is filtering on selected, debounce so users can click and unclick multiple without multiple emits or the item vanishing immediately
    if (this.selectedTab && !!this.filter[this.selectedTab]?.["_selected"]) this.debounceAndMark(true);
    else this.emitSelection();
  }

  private emitSelection(): void {
    if (!this.selectedTab) this.selected.emit({ data: [], tab: undefined });
    else {
      const selectedRows = this.rows(this.selectedTab).filter((row) => row._selected);
      this.selected.emit({ data: selectedRows, tab: this.selectedTab });
    }
    queueMicrotask(() => {
      try {
        this.chRef.detectChanges();
      } catch (error) {
        this.reportRuntimeError("emitSelection change detection", error);
        this.chRef.markForCheck();
      }
    });
  }

  // sorts, filters, column visibility changes
  private sortChange(col: SharepointChoiceColumn): void {
    if (!col.field || col.sortable === false || !this.selectedTab) return;

    // Get current sort state
    const currentSort = { ...this.sort };
    if (!currentSort[this.selectedTab]) currentSort[this.selectedTab] = [];

    const ascendingIndex = currentSort[this.selectedTab]!.findIndex(
      (sortEntry) => sortEntry.field == col.field && sortEntry.direction == "asc",
    );
    const descendingIndex = currentSort[this.selectedTab]!.findIndex(
      (sortEntry) => sortEntry.field == col.field && sortEntry.direction == "desc",
    );

    if (ascendingIndex == -1 && descendingIndex == -1) {
      currentSort[this.selectedTab] = [...currentSort[this.selectedTab]!, { field: col.field, direction: "asc" }];
    } else if (ascendingIndex >= 0) {
      currentSort[this.selectedTab] = [...currentSort[this.selectedTab]!];
      currentSort[this.selectedTab]![ascendingIndex] = {
        field: col.field,
        direction: "desc",
      };
    } else {
      currentSort[this.selectedTab] = currentSort[this.selectedTab]!.filter(
        (sortEntry) => !(sortEntry.field == col.field),
      );
    }

    // Reassign to trigger setter
    this.sort = currentSort;
  }

  private headerClick(col: SharepointChoiceColumn): void {
    if (this._suppressHeaderClick) {
      this._suppressHeaderClick = false;
      return;
    }
    this.sortChange(col);
  }

  // on multi select, not using ctrl key
  private changed(col: SharepointChoiceColumn, e: Event): void {
    e.stopPropagation();
    const target = e.target as HTMLOptionElement;

    const scrollTop = target.parentElement?.scrollTop ?? 0;

    // use inner text as value gets adjusted by angular
    const selectedText = target.innerText;
    if (selectedText === "-- All --") return this.filterChange(col, "equals", []);

    // ensure object type is correct or replace the object
    const selectedValues: string[] = [];
    const selectedTab = this.selectedTab;
    if (!selectedTab || !col.field) return;
    const currentFilterValues = this.filter[selectedTab]?.[col.field]?.equals;
    if (currentFilterValues !== undefined && currentFilterValues !== null) {
      if (Array.isArray(currentFilterValues) && currentFilterValues.every((value) => typeof value === "string")) {
        selectedValues.push(...currentFilterValues);
      } else {
        this.reportInvalidInput(
          `filter["${selectedTab}"]["${col.field}"].equals`,
          "an array of strings",
          currentFilterValues,
        );
      }
    }

    // if there are selected results set the field to add/remove the most recent click
    const selectedIndex = selectedValues.indexOf(selectedText);
    if (selectedIndex >= 0) selectedValues.splice(selectedIndex, 1);
    else selectedValues.push(selectedText);

    if (target.parentElement) {
      // manually toggle html select state as mousedown false prevents the toggle as well as prevent the unselect all
      target.selected = !target.selected;
      setTimeout(() => {
        target.parentElement ? (target.parentElement.scrollTop = scrollTop) : null;
      }, 10);
      setTimeout(() => {
        target.parentElement ? target.parentElement.focus() : null;
      }, 10);
    }

    // push the change to trigger getter
    this.filterChange(col, "equals", selectedValues);
  }

  private filterChange(
    col: SharepointChoiceColumn,
    op: SharepointChoiceFilterOperation,
    value: Array<string> | Date | string | number | null,
  ): void {
    if (!col.field || col.filter == "none" || !this.selectedTab) return;
    if (value === "-- All --") value = null;

    if (typeof value === "string") {
      value = value.trim();
      // ensure any strings are correctly converted
      if (col.filter == "date" && value) {
        const date = new Date(value);
        if (!Number.isFinite(date.getTime())) {
          this.reportInvalidInput(`filterChange("${col.field}")`, "a valid date", value);
          return;
        }
        value = date;
      }
      if (col.filter == "number" && value) {
        const number = Number(value);
        if (!Number.isFinite(number)) {
          this.reportInvalidInput(`filterChange("${col.field}")`, "a finite number", value);
          return;
        }
        value = number;
      }
    }
    if (col.filter === "date" && value instanceof Date && !Number.isFinite(value.getTime())) {
      this.reportInvalidInput(`filterChange("${col.field}")`, "a valid date", value);
      return;
    }
    if (col.filter === "number" && typeof value === "number" && !Number.isFinite(value)) {
      this.reportInvalidInput(`filterChange("${col.field}")`, "a finite number", value);
      return;
    }
    if (Array.isArray(value) && !value.every((item) => typeof item === "string")) {
      this.reportInvalidInput(`filterChange("${col.field}")`, "an array of strings", value);
      value = value.filter((item): item is string => typeof item === "string");
    }

    // Get current filter state
    const currentFilter = { ...this.filter };
    if (!currentFilter[this.selectedTab]) currentFilter[this.selectedTab] = {};

    if (
      value === undefined ||
      value === null ||
      value === "" ||
      (typeof value === "object" && "length" in value && value.length == 0)
    ) {
      if (currentFilter[this.selectedTab]![col.field]) {
        delete currentFilter[this.selectedTab]![col.field]![op];
        if (Object.keys(currentFilter[this.selectedTab]![col.field]!).length == 0)
          delete currentFilter[this.selectedTab]![col.field];
      }
    } else {
      if (!currentFilter[this.selectedTab]![col.field]) currentFilter[this.selectedTab]![col.field] = { [op]: value };
      else
        currentFilter[this.selectedTab]![col.field] = {
          ...currentFilter[this.selectedTab]![col.field],
          [op]: value,
        };
    }

    // Reassign to trigger setter
    this.filter = currentFilter;
  }

  private sortContains(field: string | undefined, direction: "asc" | "desc"): boolean {
    if (!field || !this.selectedTab || !this.sort[this.selectedTab]) return false;
    return this.sort[this.selectedTab]!.some(
      (sortEntry) => sortEntry.field == field && (direction ? sortEntry.direction == direction : true),
    );
  }

  private filterContains(field?: string) {
    const tabFilter = this.selectedTab ? this.filter[this.selectedTab] : undefined;

    // if filter exists on selected then all others disabled
    if (this.selectedTab && tabFilter?.["_selected"] && field != "_selected")
      return {
        equals: null,
        contains: null,
        greater: null,
        less: null,
      };

    return {
      equals: !field || !tabFilter ? null : (tabFilter[field]?.equals ?? null),
      contains: !field || !tabFilter ? null : (tabFilter[field]?.contains ?? null),
      greater: !field || !tabFilter ? null : (tabFilter[field]?.greater ?? null),
      less: !field || !tabFilter ? null : (tabFilter[field]?.less ?? null),
    };
  }

  private filterToggle(col: SharepointChoiceColumn, event: Event): void {
    this.allCols.forEach((column) => {
      if (column.field != col.field) column._filtervisible = false;
    });
    col._filtervisible = !col._filtervisible;
    event.stopPropagation();
    this.chRef.markForCheck();
  }

  private resetFilters(resetColumnOrder: boolean = true): void {
    this.sort = {};
    this.filter = {};
    if (resetColumnOrder) {
      this.columnOrder = {};
      this.hiddenColumns = {};
    }
  }

  private showClearButton(): boolean {
    return this.hasActiveValue(this.search) || this.hasActiveValue(this.sort) || this.hasActiveValue(this.filter);
  }

  // Search is controlled by the parent and column layout is intentionally retained.
  private clearTable(event?: Event): void {
    this.resetFilters(false);
    this.cleared.emit();
  }

  private async exportTable(currentRows: SharepointChoiceRow[]): Promise<void> {
    try {
      await this.exportTableCore(currentRows);
    } catch (error) {
      this.errorHandler.handleError(error);
    }
  }

  private async exportTableCore(currentRows: SharepointChoiceRow[]): Promise<void> {
    if (!this.export || !this.selectedTab) return;

    const columnsForExport: SharepointChoiceColumn[] = [];
    this.fields(this.selectedTab).forEach((col) => {
      if (col.field) columnsForExport.push(col);
      col.children?.forEach((child) => {
        if (child.field) columnsForExport.push(child);
      });
    });

    const baseContext = {
      tab: this.selectedTab,
      filteredRows: currentRows,
      selectedRows: currentRows.filter((row) => row._selected),
      visibleColumns: columnsForExport,
    };

    let sourceRows = baseContext.filteredRows;
    if (typeof this.export.sourceRows === "function") sourceRows = await this.export.sourceRows(baseContext);
    else if (this.export.sourceRows === "selected") sourceRows = baseContext.selectedRows;
    else if (this.export.sourceRows === "selected-or-filtered")
      sourceRows = baseContext.selectedRows.length > 0 ? baseContext.selectedRows : baseContext.filteredRows;

    const context: SharepointChoiceExportContext = {
      ...baseContext,
      sourceRows,
    };

    let columns = context.visibleColumns
      .filter((col) => !!col.field)
      .map(
        (col) =>
          ({
            header: col.headerName ?? col.field!,
            key: col.field!,
            width: col.width,
          }) as SharepointChoiceExportColumn,
      );
    if (typeof this.export.columns === "function") columns = await this.export.columns(context);
    else if (this.export.columns) columns = this.export.columns;

    let rows = context.sourceRows.map((row) => {
      const exportRow: Record<string, unknown> = {};
      columns.forEach((column) => {
        if (!column.key) return;
        exportRow[column.key] = column.key
          .split(".")
          .reduce<unknown>((value, part) => (value == null ? null : (value as Record<string, unknown>)[part]), row);
      });
      return exportRow;
    });
    if (typeof this.export.rows === "function") rows = await this.export.rows(context);
    else if (this.export.rows) rows = this.export.rows;

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet(
      (typeof this.export.sheetName === "function" ? this.export.sheetName(context) : this.export.sheetName) || "Sheet",
    );
    const title = typeof this.export.title === "function" ? this.export.title(context) : this.export.title;

    worksheet.columns = columns;
    if (title) {
      const titleValues = Array.isArray(title) ? title : [title];
      const titleRow = worksheet.insertRow(1, titleValues);
      titleRow.font = { bold: true };
    }

    rows.forEach((row) => worksheet.addRow(row));

    if (this.export.autoWidth !== false)
      worksheet.columns.forEach((column) => {
        if (!column || typeof column.eachCell !== "function") return;

        let maxLength = 2;
        column.eachCell((cell) => {
          if (!cell?.value) return;
          maxLength = Math.max(maxLength, cell.value.toString().length);
        });
        column.width = column.width && maxLength + 2 < column.width ? column.width : maxLength + 2;
      });

    const data = await workbook.xlsx.writeBuffer();
    const file = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(file);
    const link = document.createElement("a");
    link.href = url;
    link.download =
      (typeof this.export.fileName === "function" ? this.export.fileName(context) : this.export.fileName) ||
      `${context.tab}.xlsx`;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  private columnToggle(col: SharepointChoiceColumn, tab: string): void {
    const currentHidden = { ...this.hiddenColumns };
    if (!col.field) return;

    if (!currentHidden[tab]) currentHidden[tab] = [];

    if (!this.isHidden(col, tab)) currentHidden[tab] = [...currentHidden[tab], col.field];
    else currentHidden[tab] = currentHidden[tab].filter((field) => field != col.field);

    // Reassign to trigger setter
    this.hiddenColumns = currentHidden;
  }

  private toggleRow(col: SharepointChoiceColumn, checked: boolean): void {
    this.allTabs.forEach((tabName) => {
      if (checked && !this.isHidden(col, tabName)) this.columnToggle(col, tabName);
      else if (!checked && this.isHidden(col, tabName)) this.columnToggle(col, tabName);
    });
  }

  private toggleChildren(col: SharepointChoiceColumn, tab: string | null, checked: boolean): void {
    this.allTabs
      .filter((tabName) => !tab || tabName == tab)
      .forEach((tabName) => {
        if (!col.children) return;
        col.children.forEach((child) => {
          if (checked && !this.isHidden(child, tabName)) this.columnToggle(child, tabName);
          else if (!checked && this.isHidden(child, tabName)) this.columnToggle(child, tabName);
        });
      });
  }

  private isReorderable(col: SharepointChoiceColumn, parent?: SharepointChoiceColumn): boolean {
    return !!this.getColumnDragKey(col, parent);
  }

  private dragStart(event: DragEvent, col: SharepointChoiceColumn, parent?: SharepointChoiceColumn): void {
    const dragKey = this.getColumnDragKey(col, parent);
    if (!this.selectedTab || !dragKey) return;

    this._dragColumn = col;
    this._dragParent = parent;

    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.dropEffect = "move";
      event.dataTransfer.setData("text/plain", dragKey);
    }
  }

  private dragOver(event: DragEvent, col: SharepointChoiceColumn, parent?: SharepointChoiceColumn): void {
    if (!this.canDropColumn(col, parent)) return;

    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
  }

  private dragDrop(event: DragEvent, col: SharepointChoiceColumn, parent?: SharepointChoiceColumn): void {
    if (!this.canDropColumn(col, parent) || !this.selectedTab || !this._dragColumn) return;

    event.preventDefault();
    event.stopPropagation();

    const nextOrder = parent
      ? this.moveField(
          this.columnOrder[this.selectedTab] || [],
          this._dragColumn.field || "",
          col.field || "",
          (parent.children || []).filter((sibling) => !!sibling.field).map((sibling) => sibling.field as string),
        )
      : this.moveColumnGroup(
          this.columnOrder[this.selectedTab] || [],
          this._dragColumn,
          col,
          this.fields(this.selectedTab),
        );

    this.columnOrder = {
      ...this.columnOrder,
      [this.selectedTab]: nextOrder,
    };

    this._suppressHeaderClick = true;
    this.dragEnd();
  }

  private dragEnd(): void {
    this._dragColumn = undefined;
    this._dragParent = undefined;
  }

  // handle column resizing drag
  private startResize(event: MouseEvent, col: SharepointChoiceColumn): void {
    event.preventDefault();
    event.stopPropagation();
    this._resizeCleanup?.();
    const startX = event.pageX;
    const target = event.target as HTMLElement;
    const startWidth = target.parentNode ? (target.parentNode as HTMLElement).offsetWidth : col.width || 100;
    const min = parseInt(target.getAttribute("start-width") || "0");
    // The first rendered width is intentionally retained as this column's minimum width.
    if (!min) target.setAttribute("start-width", startWidth.toString());
    target.style.border = "1px solid #000";

    let onMouseUp: (e: MouseEvent) => void;
    let onWindowBlur: () => void;
    const onMouseMove = (e: MouseEvent) => {
      target.style.right = startX - e.pageX + "px";
    };
    const cleanup = () => {
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
      window.removeEventListener("blur", onWindowBlur);
      target.style.border = "none";
      target.style.right = "0px";
      if (this._resizeCleanup === cleanup) this._resizeCleanup = undefined;
    };
    onMouseUp = (mouseEvent: MouseEvent) => {
      cleanup();
      const newWidth = Math.max(startWidth + (mouseEvent.pageX - startX), min);
      col.width = newWidth;
      this.chRef.markForCheck();
    };
    onWindowBlur = () => cleanup();
    this._resizeCleanup = cleanup;
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
    window.addEventListener("blur", onWindowBlur);
  }

  // handle cell click or row click, return true or false to current cell editing, done via then to avoid await in template as it doesnt impact outcome
  private handleCellClick(
    col: SharepointChoiceColumn,
    row: SharepointChoiceRow,
    event: Event | SharepointChoiceChangeEvent,
  ): void {
    const eventTarget = event.target instanceof Element ? event.target : undefined;
    const isChoice = eventTarget?.tagName == "APP-CHOICE";
    // if its editable and not editing already (spc onchange from app-choice will return the component tag)
    if (col.spec && col.field && !isChoice) {
      // show this app-choice for editing early to later get focus
      row._editing = col.field;
      this.chRef.markForCheck();
      // get the target cell to focus after render
      let cellTarget = eventTarget;
      // ensure we are at the cell level not any inner nodes
      while (cellTarget && cellTarget.tagName != "TD") cellTarget = cellTarget.parentElement ?? undefined;
      // failed to find TD then end
      if (!cellTarget) return;
      // await render, init, all internals of app-choice then focus to edit
      // The fixed delay is intentional: nested choice controls may need time to load and initialize their options.
      setTimeout(() => {
        cellTarget.querySelector<HTMLElement>("select, textarea, input")?.focus();
      }, 500);
    } else {
      // get the trigger functions in priority order
      const clickResult = col.cellClicked
        ? col.cellClicked(row, eventTarget)
        : this.rowClicked
          ? this.rowClicked(row, eventTarget)
          : undefined;
      if (!col.cellClicked && !this.rowClicked) this.clicked.emit({ row, target: eventTarget });
      // ensure editing ends/doesnt exist but after using target above
      row._editing = undefined;
      this.chRef.markForCheck();
      // cant clear cache end
      if (!this.selectedTab) return;
      // wont need to clear cache end
      if (!isChoice && !clickResult) return;
      // if it is choice or got a function truethy outward reset cache
      if (!(clickResult instanceof Promise)) {
        this.invalidateRows(true);
        this.chRef.markForCheck();
      } else {
        clickResult
          .then((result: unknown) => {
            const selectedTab = this.selectedTab;
            if (result || isChoice) {
              if (!selectedTab) return;
              this.invalidateRows(true);
              this.chRef.detectChanges();
            }
          })
          .catch((error: unknown) => {
            this.errorHandler.handleError(error);
          });
      }
    }
  }

  private beingObserved(col: SharepointChoiceColumn): boolean {
    return !!col.spec || !!col.cellClicked || this._isObserved;
  }

  // calculate the row hyperlink only if there isnt editable, cell click, row click or clicked first
  private hyperlink(row: SharepointChoiceRow): string | undefined {
    return this.hyperlinkRow?.(row);
  }

  private sharepointChoiceField(field: string): string {
    return field.substring(field.lastIndexOf(".") + 1);
  }

  private sharepointChoiceSpec(spec: SharepointChoiceField, field: string): SharepointChoiceList {
    const fieldName = this.sharepointChoiceField(field);
    const choiceSpec: SharepointChoiceList = {};
    choiceSpec[fieldName] = {
      ...spec,
      Title: "",
    };
    return choiceSpec;
  }

  private sharepointChoiceRefreshKey(row: SharepointChoiceRow, field?: string): string {
    return `${row["_tracking"] ?? ""}:${field ?? ""}:${this._refreshCycle}`;
  }

  private sharepointChoiceForm(row: SharepointChoiceRow, field: string): SharepointChoiceForm {
    const fieldParts = this.fieldPart(field);
    let formValue: unknown = row;
    for (let index = 0; index < fieldParts.length - 1; index++) {
      formValue = (formValue as Record<string, unknown>)[fieldParts[index]!];
    }
    return formValue as SharepointChoiceForm;
  }

  private isHidden(col: SharepointChoiceColumn, tab: string): boolean {
    if (typeof col.hide == "function") return col.hide(tab);
    if (!col.field || !this.allowHideColumns) return false;
    return this.hiddenColumns[tab]?.includes(col.field) ?? false;
  }

  // get columns based on tab and hidden state
  private fields(tab: string): SharepointChoiceColumn[] {
    // Return cached result if available (cache cleared on hide changes)
    const cached = this._colsCache.get(tab);
    if (cached && cached.length > 0) return cached;

    const order = new Map<string, number>();
    (this.columnOrder[tab] || []).forEach((field, index) => {
      if (!order.has(field)) order.set(field, index);
    });

    // hide based on tab function or hide state
    const columns = this.allCols
      .filter((column) => {
        return !this.isHidden(column, tab);
      })
      .map((column) => {
        if (column.children) {
          const visibleColumn = { ...column };
          visibleColumn.children = visibleColumn.children?.filter((child) => !this.isHidden(child, tab));
          return visibleColumn;
        }
        return column;
      })
      .filter((column) => {
        if (column.children) return column.children.length > 0;
        return true;
      });

    const ordered = this.orderColumns(columns, order);

    // Cache the result
    this._colsCache.set(tab, ordered);
    return ordered;
  }

  // utility functions
  private niceName(col: SharepointChoiceColumn): string {
    if (col.headerName) return col.headerName;
    if (!col.field) return "";
    const fieldName = col.field.substring(col.field.lastIndexOf(".") + 1);
    return fieldName ? fieldName.charAt(0).toUpperCase() + fieldName.slice(1).replace(/([a-z])([A-Z])/g, "$1 $2") : "";
  }

  private fieldValue(row: SharepointChoiceRow, field?: string): SharepointChoiceRowValue {
    if (!field) return null;
    const fieldParts = this.fieldPart(field);
    let value: SharepointChoiceRowValue = row[fieldParts[0]!];
    for (let index = 1; value && index < fieldParts.length; index++) {
      if (typeof value !== "object" || value instanceof Date) return undefined;
      value = value[fieldParts[index]!];
    }
    return value;
  }

  private compareFilterValues(
    fieldValue: SharepointChoiceRowValue,
    filterValue: number | Date,
  ): number | undefined {
    const fieldNumber = fieldValue instanceof Date ? fieldValue.getTime() : Number(fieldValue);
    const filterNumber = filterValue instanceof Date ? filterValue.getTime() : Number(filterValue);
    if (Number.isNaN(fieldNumber) || Number.isNaN(filterNumber)) return undefined;
    if (fieldNumber < filterNumber) return -1;
    if (fieldNumber > filterNumber) return 1;
    return 0;
  }

  private compareRowValues(leftValue: SharepointChoiceRowValue, rightValue: SharepointChoiceRowValue): number {
    const toComparableValue = (value: SharepointChoiceRowValue): string | number | boolean | null | undefined => {
      if (value instanceof Date) return value.getTime();
      if (typeof value === "object" && value !== null) return String(value);
      return value;
    };
    const left = toComparableValue(leftValue);
    const right = toComparableValue(rightValue);

    if (typeof left === "string" && typeof right === "string") {
      if (left < right) return -1;
      if (left > right) return 1;
      return 0;
    }

    const leftNumber = Number(left);
    const rightNumber = Number(right);
    if (Number.isNaN(leftNumber) || Number.isNaN(rightNumber)) return 0;
    if (leftNumber < rightNumber) return -1;
    if (leftNumber > rightNumber) return 1;
    return 0;
  }

  private fieldPart(field: string): Array<string> {
    const cached = this._fieldMapCache.get(field);
    if (cached && cached.length > 0) return cached;
    const parts = field.split(".");
    this._fieldMapCache.set(field, parts);
    return parts;
  }

  private canDropColumn(col: SharepointChoiceColumn, parent?: SharepointChoiceColumn): boolean {
    if (!this.isReorderable(col, parent) || !this._dragColumn) return false;
    const draggedKey = this.getColumnDragKey(this._dragColumn, this._dragParent);
    const targetKey = this.getColumnDragKey(col, parent);
    if (!draggedKey || !targetKey) return false;
    if (draggedKey == targetKey) return false;
    if (!this._dragParent && !parent) return true;
    return this._dragParent === parent;
  }

  private getColumnDragKey(col: SharepointChoiceColumn, parent?: SharepointChoiceColumn): string | undefined {
    if (col.field) return col.field;
    if (parent || !col.children?.length) return undefined;

    const explicitName = col.headerName?.trim();
    if (explicitName) return `__group__:${explicitName}`;

    const childKey = col.children
      .map((child) => child.field || child.headerName || "")
      .filter((value) => !!value)
      .join("|");
    return childKey ? `__group__:${childKey}` : undefined;
  }

  private getColumnFields(col: SharepointChoiceColumn): string[] {
    if (col.field) return [col.field];
    return col.children?.map((child) => child.field).filter((field): field is string => !!field) || [];
  }

  private moveColumnGroup(
    order: string[],
    draggedColumn: SharepointChoiceColumn,
    targetColumn: SharepointChoiceColumn,
    siblings: SharepointChoiceColumn[],
  ): string[] {
    const nextOrder = order.filter((field, index, values) => !!field && values.indexOf(field) == index);
    const siblingFields = siblings.flatMap((col) => this.getColumnFields(col));
    siblingFields.forEach((field) => {
      if (!nextOrder.includes(field)) nextOrder.push(field);
    });

    const draggedFields = this.getColumnFields(draggedColumn);
    const targetFields = this.getColumnFields(targetColumn);
    if (draggedFields.length == 0 || targetFields.length == 0) return nextOrder;

    const remaining = nextOrder.filter((field) => !draggedFields.includes(field));
    const targetField = targetFields.find((field) => remaining.includes(field));
    if (!targetField) return nextOrder;

    const targetIndex = remaining.indexOf(targetField);
    remaining.splice(targetIndex, 0, ...draggedFields.filter((field) => siblingFields.includes(field)));
    return remaining;
  }

  private moveField(order: string[], draggedField: string, targetField: string, siblings: string[]): string[] {
    const nextOrder = order.filter((field, index, values) => !!field && values.indexOf(field) == index);

    siblings.forEach((field) => {
      if (!nextOrder.includes(field)) nextOrder.push(field);
    });

    const from = nextOrder.indexOf(draggedField);
    const to = nextOrder.indexOf(targetField);

    if (from < 0 || to < 0) return nextOrder;

    nextOrder.splice(from, 1);
    nextOrder.splice(from < to ? to - 1 : to, 0, draggedField);
    return nextOrder;
  }

  private orderColumns(columns: SharepointChoiceColumn[], order: Map<string, number>): SharepointChoiceColumn[] {
    const ordered = this.sortColumns(columns, order);

    return ordered.map((col) => {
      if (!col.children || col.children.length < 2) return col;

      return {
        ...col,
        children: this.sortColumns(col.children, order, col),
      };
    });
  }

  private sortColumns(
    columns: SharepointChoiceColumn[],
    order: Map<string, number>,
    parent?: SharepointChoiceColumn,
  ): SharepointChoiceColumn[] {
    if (columns.length < 2 || order.size == 0) return columns;

    const reorderable = columns
      .map((col, index) => ({
        col,
        index,
        sortIndex: this.getColumnSortIndex(col, order, parent),
      }))
      .filter((item) => item.sortIndex < Number.MAX_SAFE_INTEGER);
    if (reorderable.length < 2) return columns;

    const sorted = [...reorderable]
      .sort((leftItem, rightItem) => {
        if (leftItem.sortIndex == rightItem.sortIndex) return leftItem.index - rightItem.index;
        return leftItem.sortIndex - rightItem.sortIndex;
      })
      .map((item) => item.col);

    let position = 0;
    return columns.map((col) => {
      if (this.getColumnSortIndex(col, order, parent) >= Number.MAX_SAFE_INTEGER) return col;

      return sorted[position++] ?? col;
    });
  }

  private getColumnSortIndex(
    col: SharepointChoiceColumn,
    order: Map<string, number>,
    parent?: SharepointChoiceColumn,
  ): number {
    if (parent) return col.field && order.has(col.field) ? order.get(col.field)! : Number.MAX_SAFE_INTEGER;

    const indices = this.getColumnFields(col)
      .filter((field) => order.has(field))
      .map((field) => order.get(field)!);
    return indices.length > 0 ? Math.min(...indices) : Number.MAX_SAFE_INTEGER;
  }

  private hasActiveValue(value: unknown): boolean {
    if (value === null || value === undefined || value === "") return false;
    if (value instanceof Date) return !isNaN(value.getTime());
    if (Array.isArray(value)) return value.some((item) => this.hasActiveValue(item));
    if (typeof value === "object") return Object.values(value).some((item) => this.hasActiveValue(item));
    return true;
  }

  private filterDistinct(field?: string, split: boolean = false): SharepointChoiceDistinctValue[] {
    if (!field || !this.allData || !this.selectedTab) return [];

    const tab = this.selectedTab;
    const cacheKey = JSON.stringify([tab, field, split]);
    if (this._distinctCache.has(cacheKey)) return this._distinctCache.get(cacheKey)!;

    const tabData = this.allData[tab];
    if (!tabData) return [];

    const values: SharepointChoiceDistinctValue[] = [];
    tabData.forEach((row) => {
      const cellValue = this.fieldValue(row, field);
      if ((cellValue === null || cellValue === undefined || cellValue === "") && !values.includes("(blanks)")) {
        values.push("(blanks)");
      } else if (cellValue === 0 || cellValue === false || cellValue) {
        // if filter is split then split the value by comma or semicolon and trim to get distinct values for multi choice fields
        if (split && typeof cellValue === "string" && (cellValue.includes(",") || cellValue.includes(";"))) {
          cellValue
            .split(/[,;]/)
            .map((value) => value.trim())
            .forEach((value) => {
              if (!values.includes(value)) values.push(value);
            });
        } else if (!values.includes(cellValue)) {
          values.push(cellValue);
        }
      }
    });
    const sorted = values.sort((leftValue, rightValue) => {
      if (leftValue === "(blanks)") return -1;
      if (rightValue === "(blanks)") return 1;
      const left =
        leftValue instanceof Date ? leftValue.getTime() : typeof leftValue === "object" ? String(leftValue) : leftValue;
      const right =
        rightValue instanceof Date
          ? rightValue.getTime()
          : typeof rightValue === "object"
            ? String(rightValue)
            : rightValue;
      if (typeof left === "string" && typeof right === "string") {
        if (left < right) return -1;
        if (left > right) return 1;
        return 0;
      }
      const leftNumber = Number(left);
      const rightNumber = Number(right);
      if (Number.isNaN(leftNumber) || Number.isNaN(rightNumber)) return 0;
      if (leftNumber < rightNumber) return -1;
      if (leftNumber > rightNumber) return 1;
      return 0;
    });
    this._distinctCache.set(cacheKey, sorted);
    return sorted;
  }

  private currentPageRows(rows: SharepointChoiceRow[]): SharepointChoiceRow[] {
    const cached = this._pageCache;
    if (cached && cached.length > 0) return cached;

    const result = rows.slice((this.pageNumber - 1) * this.pageSize, this.pageNumber * this.pageSize);
    this._pageCache = result;

    return result;
  }

  private trackRow(index: number, row: SharepointChoiceRow): string | number {
    return row._tracking ?? index;
  }

  // get rows based on tab, sort and filter
  private rows(tab: string): SharepointChoiceRow[] {
    // Return cached result if available (cache cleared on filter/sort changes)
    if (this._rowsCache.has(tab)) return this._rowsCache.get(tab)!;

    this._pageCache = [];
    this._isObserved = !!this.rowClicked || this.clicked.observed;
    this._nodeCache.clear();

    const sourceFilters = this.filter[tab] || {};
    let activeFilters: SharepointChoiceFilter[string] = {};

    for (const fieldName in sourceFilters) {
      const fieldFilter = sourceFilters[fieldName];
      if (!fieldFilter) continue;

      activeFilters[fieldName] = { ...fieldFilter };
      if (typeof fieldFilter.greater === "string") activeFilters[fieldName].greater = new Date(fieldFilter.greater);
      if (typeof fieldFilter.less === "string") activeFilters[fieldName].less = new Date(fieldFilter.less);
    }

    const sortEntries = this.sort[tab] || [];

    if (activeFilters["_selected"]) activeFilters = { _selected: activeFilters["_selected"] };

    const hasFilter = Object.keys(activeFilters).some((fieldName) => {
      if (this.isHidden({ field: fieldName }, tab)) return false;
      return true;
    });

    const hasSort =
      sortEntries.length > 0 &&
      sortEntries.some((sortEntry) => {
        if (this.isHidden({ field: sortEntry.field }, tab)) return false;
        return true;
      });

    // never mutate the input array (this.allData[tab]) if there is sorting
    let resultRows: SharepointChoiceRow[] = hasSort ? [...(this.allData[tab] ?? [])] : (this.allData[tab] ?? []);

    if (hasFilter) {
      resultRows = resultRows.filter((row) => {
        // apply all filters
        for (const fieldName in activeFilters) {
          // user hidden columns to skip filters
          if (this.isHidden({ field: fieldName }, tab)) continue;

          const operations = activeFilters[fieldName];
          if (!operations) continue;
          for (const operation of ["contains", "equals", "greater", "less"] as const) {
            const filterValue = operations[operation];
            if (filterValue === null || filterValue === undefined) continue;
            // get the field value
            const cellValue = this.fieldValue(row, fieldName);
            // apply the operation, all inverted as its easier to think about what is required
            if (operation == "contains") {
              if (!cellValue?.toString()?.toLowerCase()?.includes(filterValue.toString().toLowerCase())) return false;
            } else if (operation == "equals") {
              if (Array.isArray(filterValue)) {
                if (
                  !filterValue.some(
                    (selectedValue: string) =>
                      cellValue?.toString() === selectedValue.toString() ||
                      (selectedValue == "(blanks)" &&
                        (cellValue === null || cellValue === undefined || cellValue === "")),
                  )
                )
                  return false;
              } else if (
                !(
                  cellValue?.toString() === filterValue.toString() ||
                  (filterValue == "(blanks)" && (cellValue === null || cellValue === undefined || cellValue === ""))
                )
              )
                return false;
            } else if (operation == "greater" || operation == "less") {
              if (cellValue === null || cellValue === undefined || cellValue === "") return false;
              const comparisonValue = operation === "greater" ? operations.greater : operations.less;
              if (comparisonValue === null || comparisonValue === undefined) continue;
              const comparison = this.compareFilterValues(cellValue, comparisonValue);
              if (comparison === undefined) return false;
              if (operation == "greater" && comparison < 0) return false;
              if (operation == "less" && comparison > 0) return false;
            }
          }
        }
        return true;
      });
    }

    if (hasSort)
      resultRows.sort((leftRow, rightRow) => {
        for (const sortEntry of sortEntries) {
          // user hidden columns to skip sorts
          if (this.isHidden({ field: sortEntry.field }, tab)) continue;

          const leftValue = this.fieldValue(leftRow, sortEntry.field);
          const rightValue = this.fieldValue(rightRow, sortEntry.field);
          const direction = sortEntry.direction == "desc" ? -1 : 1;

          if (leftValue === null && rightValue === null) continue;
          if (leftValue === null) return 1;
          if (rightValue === null) return -1;
          const comparison = this.compareRowValues(leftValue, rightValue);
          if (comparison !== 0) return comparison * direction;
        }

        return 0;
      });

    // Cache the result
    this._rowsCache.set(tab, resultRows);
    return resultRows;
  }

  private renderNode(
    col: SharepointChoiceColumn,
    value: SharepointChoiceRowValue,
    row: SharepointChoiceRow,
    rowIndex: number,
    colIndex: number,
    childIndex?: number,
  ): SharepointChoiceRenderContent {
    if (!col.cellRenderer) return null;

    // could return issues if cols hidden and many aren't named at all
    const key = JSON.stringify([
      this.selectedTab ?? "",
      row["_tracking"] ?? rowIndex,
      col.field ?? col.headerName ?? (1000 * colIndex + (childIndex ?? 0)).toString(),
    ]);

    if (!this._nodeCache.has(key)) {
      this._nodeCache.set(key, col.cellRenderer(value, row, rowIndex));
    }
    const content = this._nodeCache.get(key);

    const hyperlink = this.beingObserved(col) ? undefined : this.hyperlink(row);
    if (!hyperlink) return content;

    return {
      content,
      hyperlink,
      hyperlinkTarget: this.hyperlinkTarget,
    };
  }

  // Helper method to determine cell type (reduces template complexity)
  private getCellType(
    col: SharepointChoiceColumn,
    value: SharepointChoiceRowValue,
  ): "renderer" | "date" | "number" | "boolean" | "default" {
    // quick checks for type
    if (col.cellRenderer) return "renderer";
    if (col.filter === "date") return "date";
    if (col.filter === "number") return "number";
    // slower checks based on value
    switch (typeof value) {
      case "number":
        return "number";
      case "boolean":
        return "boolean";
    }
    // check for date object
    return value instanceof Date ? "date" : "default";
  }

  // Helper to format boolean values
  private formatBoolean(value: SharepointChoiceRowValue): string {
    return value === true ? "✔" : value === false ? "✘" : "";
  }

  private datePipeValue(value: SharepointChoiceRowValue): string | number | Date | null | undefined {
    return typeof value === "string" || typeof value === "number" || value instanceof Date ? value : null;
  }

  private numberPipeValue(value: SharepointChoiceRowValue): string | number | null | undefined {
    return typeof value === "string" || typeof value === "number" ? value : null;
  }

  private ceil(number: number): number {
    return Math.ceil(number);
  }
}
