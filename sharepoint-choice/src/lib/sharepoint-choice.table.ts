import { Component, Input, ErrorHandler, EventEmitter, Output, ChangeDetectorRef, ChangeDetectionStrategy, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SharepointChoiceColumn, SharepointChoiceTabs, SharepointChoiceRow, SharepointChoiceForm, SharepointChoiceField, SharepointChoiceList, ExcelIcon } from './sharepoint-choice.models';
import { SharepointChoiceComponent } from './sharepoint-choice.component';
import { SharepointChoiceRender } from './sharepoint-choice.render';

interface SharepointChoiceHide {
  [tabName: string]: string[];
}

interface SharepointChoiceSort {
  [tabName: string]: {
    direction: 'asc' | 'desc';
    field: string;
  }[];
}

interface SharepointChoiceFilter {
  [tabName: string]: {
    [fieldName:string]: {
      equals?: string | number | boolean | Date | null;
      contains?: string | null;
      greater?: number | Date | null;
      less?: number | Date | null;
    };
  }
}

@Component({
  selector: 'app-table',
  templateUrl: './sharepoint-choice.table.html',
  styleUrls: ['../styles.scss'],
  standalone: true,
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [
    CommonModule,
    FormsModule,
    SharepointChoiceComponent,
    SharepointChoiceRender
  ],
  providers: [{
    provide: ErrorHandler
  }]
})
export class SharepointChoiceTable implements OnInit, OnDestroy {
  // all data passed in keyed by tab name (not via signals which reduce performance on large data sets)
  @Input() set allData(value: SharepointChoiceTabs) {
    this._allData = value;
    this.computedData(true);
    // Clear cache when data changes
    this._rowsCache.clear();
    this.chRef.markForCheck();
  }
  get allData(): SharepointChoiceTabs {
    return this._allData || {};
  }
  private _allData: SharepointChoiceTabs = {};

  // all columns
  @Input() set allCols(value: SharepointChoiceColumn[]) {
    this._allCols = value;
    // Clear cache when columns change
    this._colsCache.clear();
    this._nodeCache.clear();
    this._pageCache = [];
    this.chRef.markForCheck();
  }
  get allCols(): SharepointChoiceColumn[] {
    return this._allCols || [];
  }
  private _allCols: SharepointChoiceColumn[] = [];

  // all tabs or derive from all data
  @Input() set allTabs(value: string[]) {
    this._allTabs = value;
    this.computedData();
    this.chRef.markForCheck();
  }
  get allTabs(): string[] {
    return this._allTabs;
  }
  private _allTabs: string[] = [];

  // selected tab else use stored value or first tab
  @Input() set selectedTab(value: string | undefined) {
    this._selectedTab = value;
    this.computedData();
    this.setStorage(`Tab`, this._selectedTab);
    // Revert to page 1 on every tab change
    this.pageNumber = 1;
    this.chRef.markForCheck();
  }
  get selectedTab(): string | undefined {
    return this._selectedTab;
  }
  private _selectedTab?: string;

  // page size via getter/setter to allow storage
  @Input() set pageSize(value: number) {
    this._pageSize = value;
    this.setStorage(`Size`, this.pageSize);
    this._pageCache = [];
    this.chRef.markForCheck();
  }
  get pageSize(): number {
    return this._pageSize || 100;
  }
  private _pageSize?: number;

  // should show loading state and disable clicks etc
  @Input() set loading(value: boolean) {
    this._loading = value;
    this.chRef.markForCheck();
  }
  get loading(): boolean {
    return this._loading || false;
  }
  private _loading: boolean = false;

  // simple inputs that dont need getter/setter
  @Input() prefix: string = document.location.href.toLowerCase().split('?')[0].split('#')[0];
  @Input() tableHeight: string = 'calc(100vh - 360px)';
  @Input() showSingleTab: boolean = true;
  @Input() allEditing: boolean = false; // all with spec render as app-choice else edit per cell on click
  @Input() allowHideColumns: boolean = true;

  // allow selection and emit selected items and tab
  @Input() allowSelection: boolean = false;
  @Output() selected = new EventEmitter<{ data: SharepointChoiceRow[], tab: string | undefined }>();

  // outbound events or pseudo callbacks for await support on click, in order of trigger
  // cellClicked triggers ahead of all of these if present
  @Input() rowClicked?: Function; // [rowClicked]="rowClicked" rowClicked = async (rowData: any, target: HTMLElement | EventTarget | undefined) => { ... } to ensure this. is from the app and not from app-table
  @Output() clicked = new EventEmitter<{ row: SharepointChoiceRow, target: HTMLElement | EventTarget | undefined }>(); // (clicked)="onClicked($event)" onClicked(event: { row: SharepointChoiceRow, target: HTMLElement | EventTarget | undefined }) { ... }
  @Input() hyperlinkRow?: Function; // [hyperlinkRow]="hyperlinkRow" hyperlinkRow = (rowData: any) => { return 'https://...'; } to ensure this. is from the app and not from app-table
  @Input() export?: Function; // [export]="export" export = (selectedTab: string, filteredRows: SharepointChoiceRow[]) => { ... } function triggered when export icon clicked giving current tab and filtered rows

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
    this._sort = value;
    this.setStorage(`Sort`, this.sort);
    this.debounceAndMark();
  }
  get sort(): SharepointChoiceSort {
    return this._sort || {};
  }
  private _sort?: SharepointChoiceSort;

  set filter(value: SharepointChoiceFilter) {
    this._filter = value;
    this.setStorage(`Filter`, this.filter);
    this.debounceAndMark();
  }
  get filter(): SharepointChoiceFilter {
    return this._filter || {};
  }
  private _filter?: SharepointChoiceFilter;

  set hiddenColumns(value: SharepointChoiceHide) {
    this._hiddenColumns = value;
    this.setStorage(`Hide`, this.hiddenColumns);
    // Clear cache after adjusting columns which impact filters/sorts
    if (!this._selectedTab)
      return;
    this._colsCache.delete(this._selectedTab);
    this.chRef.markForCheck();
  }
  get hiddenColumns(): SharepointChoiceHide {
    return this._hiddenColumns || {};
  }
  private _hiddenColumns?: SharepointChoiceHide;

  private _dataLoadCycles: number = 0;
  private _debounceFilterSort?: ReturnType<typeof setTimeout>;
  
  // Memoization cache for columns filtered/sorted rows
  private _colsCache: Map<string, SharepointChoiceColumn[]> = new Map();
  private _rowsCache: Map<string, SharepointChoiceRow[]> = new Map();
  private _pageCache: SharepointChoiceRow[] = [];
  private _fieldMapCache: Map<string, Array<string>> = new Map();
  private _nodeCache: Map<string, any> = new Map();
  private _isObserved: boolean = false;
  
  constructor(private chRef: ChangeDetectorRef) { }

  ngOnInit(): void {
    if (!this._sort)
      try {
        this._sort = this.getStorage(`Sort`);
      } catch {}
    if (!this._filter)
      try {
        this._filter = this.getStorage(`Filter`);
      } catch {}
    if (!this._hiddenColumns)
      try {
        this._hiddenColumns = this.getStorage(`Hide`);
      } catch {}
    if (!this._selectedTab)
      try {
        this._selectedTab = this.getStorage(`Tab`);
      } catch {}
    if (!this._pageSize)
      try {
        this._pageSize = this.getStorage(`Size`);
      } catch {}
  }

  ngOnDestroy(): void {
    clearTimeout(this._debounceFilterSort);
  }

  private debounceAndMark(): void {
    // use debounce to avoid multiple rapid changes clearing cache too often across multiple columns and filter/sort types etc so simpler than rxjs
    clearTimeout(this._debounceFilterSort);
    this._debounceFilterSort = setTimeout(() => {
      if (!this.selectedTab)
        return;
      // drop cache to recalc rows
      this._rowsCache.delete(this.selectedTab);
      // emit the new filtered/sorted selection
      this.emitSelection();
    }, 500);
  }

  private computedData(incKeys:boolean = false): void {
    var tabs = Object.keys(this._allData || {}).filter(k => k && k != 'undefined' && k != 'null');

    if ((!this._allTabs || this._allTabs.length == 0) ||
      (tabs.length > 0 && this._allTabs.filter(t => tabs.includes(t)).length == 0)) {
      this._allTabs = tabs;
    }

    if (this._allTabs.length > 0
      && (!this._selectedTab || !this._allTabs.includes(this._selectedTab))) {
      this._selectedTab = this._allTabs[0];
    }

    if (!incKeys)
      return;

    this._dataLoadCycles++;
    // add _trackKey to each row for ngFor tracking
    for (let tab of tabs) {
      if (this._allData[tab]) {
        this._allData[tab].forEach((row, index) => {
          if (row['_tracking'])
            return;
          
          if (row['Id']) // likely SPId so retain through changes
            row['_tracking'] = `id-${row['Id']}`;
          else if (row['reference']) // likely policy reference so retain through changes
            row['_tracking'] = `ref-${row['reference']}`;
          else // generate unique tracking key each time data changes too
            row['_tracking'] = `${tab}-${this._dataLoadCycles}-${index}`;
        });
      }
    }
  }

  getStorage(key: string): any {
    return JSON.parse(localStorage.getItem(`SharepointTable-${this.prefix}`) || '{}')[key];
  }

  setStorage(key: string, value: any) {
    var s = JSON.parse(localStorage.getItem(`SharepointTable-${this.prefix}`) || '{}');
    s[key] = value;
    localStorage.setItem(`SharepointTable-${this.prefix}`, JSON.stringify(s));
  }

  // tab change resets page number and emits empty selection
  tabChange(tab: string): void {
    this.selectedTab = tab;
    this.emitSelection();
  }

  // selection change toggles row selected state and emits selected rows
  selectionChanged(row: SharepointChoiceRow): void {
    row._selected = !row._selected;
    // if there is filtering on selected, debounce so users can click and unclick multiple without multiple emits or the item vanishing immediately
    if (this.selectedTab && !!(this.filter[this.selectedTab]?.['_selected']))
      this.debounceAndMark();
    else
      this.emitSelection();
  }

  emitSelection(): void {
    if (!this.selectedTab)
      this.selected.emit({ data: [], tab: undefined });
    else {
      let selectedRows = this.rows(this.selectedTab).filter(r => r._selected);
      this.selected.emit({ data: selectedRows, tab: this.selectedTab });
    }
    queueMicrotask(() => {
      this.chRef.detectChanges();
    });
  }

  // sorts, filters, column visibility changes
  sortChange(col: SharepointChoiceColumn): void {
    if (!col.field || col.sortable === false || !this.selectedTab)
      return;

    // Get current sort state
    let currentSort = { ...this.sort };
    if (!currentSort[this.selectedTab])
      currentSort[this.selectedTab] = [];

    var asc = currentSort[this.selectedTab].findIndex(s => s.field == col.field && s.direction == 'asc');
    var desc = currentSort[this.selectedTab].findIndex(s => s.field == col.field && s.direction == 'desc');

    if (asc == -1 && desc == -1) {
      currentSort[this.selectedTab] = [...currentSort[this.selectedTab], { field: col.field, direction: 'asc' }];
    } else if (asc >= 0) {
      currentSort[this.selectedTab] = [...currentSort[this.selectedTab]];
      currentSort[this.selectedTab][asc] = { field: col.field, direction: 'desc' };
    } else {
      currentSort[this.selectedTab] = currentSort[this.selectedTab].filter(s => !(s.field == col.field));
    }

    // Reassign to trigger setter
    this.sort = currentSort;
  }

  filterChange(col: SharepointChoiceColumn, op: string, value: Date | string | number | null): void {
    if (!col.field || col.filter == 'none' || !this.selectedTab)
      return;

    if (typeof value === 'string') {
      value = value.trim();
      // ensure any strings are correctly converted
      if (col.filter == 'date' && value) {
        value = new Date(value);
      }
      if (col.filter == 'number' && value) {
        value = Number(value);
      }
    }

    // Get current filter state
    let currentFilter = { ...this.filter };
    if (!currentFilter[this.selectedTab])
      currentFilter[this.selectedTab] = {};

    if (value === undefined || value === null || value === '') {
      if (currentFilter[this.selectedTab][col.field]) {
        delete currentFilter[this.selectedTab][col.field][op];
        if (Object.keys(currentFilter[this.selectedTab][col.field]).length == 0)
          delete currentFilter[this.selectedTab][col.field];
      }
    } else {
      if (!currentFilter[this.selectedTab][col.field])
        currentFilter[this.selectedTab][col.field] = { [op]: value };
      else
        currentFilter[this.selectedTab][col.field] = { ...currentFilter[this.selectedTab][col.field], [op]: value };
    }

    // Reassign to trigger setter
    this.filter = currentFilter;
  }

  sortContains(field: string | undefined, direction: 'asc' | 'desc'): boolean {
    if (!field || !this.selectedTab || !this.sort[this.selectedTab])
      return false;
    return this.sort[this.selectedTab].some(s => s.field == field && (direction ? s.direction == direction : true));
  }

  filterContains(field?: string) {
    // if filter exists on selected then all others disabled
    if (this.selectedTab && this.filter[this.selectedTab]?.['_selected'] && field != '_selected')
      return {
        equals: null,
        contains: null,
        greater: null,
        less: null
      };

    return {
      equals: !field || !this.selectedTab || !this.filter[this.selectedTab] ? null : this.filter[this.selectedTab][field]?.equals ?? null,
      contains: !field || !this.selectedTab || !this.filter[this.selectedTab] ? null : this.filter[this.selectedTab][field]?.contains ?? null,
      greater: !field || !this.selectedTab || !this.filter[this.selectedTab] ? null : this.filter[this.selectedTab][field]?.greater ?? null,
      less: !field || !this.selectedTab || !this.filter[this.selectedTab] ? null : this.filter[this.selectedTab][field]?.less ?? null
    };
  }

  filterToggle(col: SharepointChoiceColumn, event: Event): void {
    this.allCols.forEach(c => {
      if (c.field != col.field)
        c._filtervisible = false;
    });
    col._filtervisible = !col._filtervisible;
    event.stopPropagation();
    this.chRef.markForCheck();
  }

  columnToggle(col: SharepointChoiceColumn, tab: string): void {
    if (!col.field)
      return;

    let currentHidden = { ...this.hiddenColumns };
    if (!currentHidden[tab])
      currentHidden[tab] = [];

    if (!this.isHidden(col, tab))
      currentHidden[tab] = [...currentHidden[tab], col.field];
    else
      currentHidden[tab] = currentHidden[tab].filter((c: string) => c != col.field);

    // Reassign to trigger setter
    this.hiddenColumns = currentHidden;
  }

  // handle column resizing drag
  startResize(event: MouseEvent, col: SharepointChoiceColumn): void {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.pageX;
    const startWidth = col.width || 100;
    var target = event.target as HTMLElement;
    target.style.border = '1px solid #000';

    const onMouseMove = (e: MouseEvent) => {
      target.style.right = startX - e.pageX + 'px';
    }
    const onMouseUp = (e: MouseEvent) => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      target.style.border = 'none';
      const newWidth = startWidth + (e.pageX - startX);
      col.width = newWidth;
      target.style.right = '0px';
      this.chRef.markForCheck();
    }
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }

  // handle cell click or row click, return true or false to current cell editing, done via then to avoid await in template as it doesnt impact outcome
  handleCellClick(col: SharepointChoiceColumn, row: SharepointChoiceRow, event: any): void {
    let choice = event.target.tagName == 'APP-CHOICE';
    // if its editable and not editing already (spc onchange from app-choice will return the component tag)
    if (col.spec && col.field && !choice) {
      // show this app-choice for editing early to later get focus
      row._editing = col.field;
      this.chRef.markForCheck();
      // get the target cell to focus after render
      var target = event.target;
      // ensure we are at the cell level not any inner nodes
      while (target && target.tagName != 'TD')
        target = target.parentNode;
      // failed to find TD then end
      if (!target)
        return;
      // await render, init, all internals of app-choice then focus to edit
      setTimeout(() => {
        var el = target.getElementsByTagName('select');
        if (el.length == 0)
          el = target.getElementsByTagName('textarea');
        if (el.length == 0)
          el = target.getElementsByTagName('input');
        if (el.length > 0)
          el[0].focus();
      }, 500);
    } else {
      // get the trigger functions in priority order
      var c: any = null;
      if (col.cellClicked)
        c = col.cellClicked(row, event.target);
      else if (this.rowClicked)
        c = this.rowClicked(row, event.target);
      else
        this.clicked.emit({ row: row, target: event.target });
      // cant clear cache end
      if (!this.selectedTab)
        return;
      // wont need to clear cache end
      if (!choice && !c)
        return;
      // if it is choice or got a function truethy outward reset cache
      if (!(c instanceof Promise))
        this._rowsCache.delete(this.selectedTab);
      else {
        c.then((r?: any) => {
          if (r || choice) {
            this._rowsCache.delete(this.selectedTab!);
            this.chRef.detectChanges();
          }
        });
      }
      // ensure editing ends/doesnt exist but after using target above
      row._editing = undefined;
      this.chRef.markForCheck();
    }
  }

  beingObserved(col: SharepointChoiceColumn): boolean {
    return !!col.spec || !!col.cellClicked || this._isObserved;
  }

  // calculate the row hyperlink only if there isnt editable, cell click, row click or clicked first
  hyperlink(row: SharepointChoiceRow): string | undefined {
    return this.hyperlinkRow?.(row);
  }

  sharepointChoiceField(field: string): string {
    return field.substring(field.lastIndexOf('.') + 1);
  }

  sharepointChoiceSpec(spec: SharepointChoiceField, field: string): SharepointChoiceList {
    var f = this.sharepointChoiceField(field);
    var s: SharepointChoiceList = {};
    s[f] = spec;
    // ensure no title to avoid label rendering
    s[f].Title = '';
    return s;
  }

  sharepointChoiceForm(row: SharepointChoiceRow, field: string): SharepointChoiceForm {
    var s = this.fieldPart(field);
    var f: any = row;
    for (let i = 0; i < s.length - 1; i++) {
      f = f[s[i]];
    }
    return f as SharepointChoiceForm;
  }

  isHidden(col: SharepointChoiceColumn, tab: string): boolean {
    if (typeof col.hide == 'function' && col.hide(tab))
      return true;
    if (!col.field || !this.hiddenColumns[tab])
      return false;
    return this.hiddenColumns[tab]?.includes(col.field);
  }

  // get columns based on tab and hidden state
  fields(tab: string): SharepointChoiceColumn[] {
    // Return cached result if available (cache cleared on hide changes)
    const cached = this._colsCache.get(tab);
    if (cached)
      return cached;

    this._rowsCache.delete(tab);

    let cols = this.allowHideColumns ? this.hiddenColumns[tab] || [] : [];
    // hide based on tab function or hide state
    const columns = this.allCols.filter(c => {
      return !cols.includes(c.field ?? '') && (!c.hide || (typeof c.hide == 'function' && !c.hide(tab)));
    }).map(c => {
      if (c.children) {
        let nc = { ...c };
        nc.children = nc.children?.filter(ch => !cols.includes(ch.field ?? '') && (!ch.hide || (typeof ch.hide == 'function' && !ch.hide(tab))));
        return nc;
      }
      return c;
    }).filter(c => {
      if (c.children)
        return c.children.length > 0;
      return true;
    });

    // Cache the result
    this._colsCache.set(tab, columns);
    return columns;
  }

  // utility functions
  niceName(col: SharepointChoiceColumn): string {
    if (col.headerName)
      return col.headerName;
    if (!col.field)
      return '';
    var h = col.field.substring(col.field.lastIndexOf('.') + 1);
    if (h)
      col.headerName = h.charAt(0).toUpperCase() + h.slice(1).replace(/([a-z])([A-Z])/g, '$1 $2');
    return col.headerName || '';
  }

  fieldValue(row: SharepointChoiceRow, field?: string): any {
    if (!field)
      return null;
    var f = this.fieldPart(field);
    var c = row[f[0]];
    for (let i = 1; c && i < f.length; i++)
      c = c[f[i]];
    return c;
  }

  private fieldPart(field: string): Array<string> {
    const cached = this._fieldMapCache.get(field);
    if (cached)
      return cached;
    const parts = field.split('.');
    this._fieldMapCache.set(field, parts);
    return parts;
  }

  filterDistinct(field?: string): any[] {
    if (!field || !this.allData || !this.selectedTab || !this.allData[this.selectedTab])
      return [];
    let values: any[] = [];
    this.allData[this.selectedTab].forEach(d => {
      var c = this.fieldValue(d, field);
      if ((c === null || c === undefined || c === '') && !values.includes('(blanks)'))
        values.push('(blanks)');
      else if ((c === 0 || c === false || c) && !values.includes(c))
        values.push(c);
    });
    return values;
  }

  currentPageRows(rows: SharepointChoiceRow[]): SharepointChoiceRow[] {
    const cached = this._pageCache;
    if (cached)
      return cached;

    const result = rows.slice((this.pageNumber - 1) * this.pageSize, this.pageNumber * this.pageSize);
    this._pageCache = result;

    return result;
  }

  // get rows based on tab, sort and filter
  rows(tab: string): SharepointChoiceRow[] {
    // Return cached result if available (cache cleared on filter/sort changes)
    const cached = this._rowsCache.get(tab);
    if (cached)
      return cached;

    this._pageCache = [];
    this._isObserved = !!this.rowClicked || this.clicked.observed;
    this._nodeCache.clear();

    var filter = this.filter[tab] || {};
    var sort = this.sort[tab] || [];

    if (filter['_selected'])
      filter = { '_selected': filter['_selected'] };

    const hasFilter = Object.keys(filter).some(field => {
      if (this.allowHideColumns && this.hiddenColumns[tab]?.includes(field))
        return false;
      return true;
    });

    const hasSort = sort && sort.length > 0 && sort.some(s => {
      if (this.allowHideColumns && this.hiddenColumns[tab]?.includes(s.field))
        return false;
      return true;
    });

    // never mutate the input array (this.allData[tab]) if there is sorting
    let result: SharepointChoiceRow[] = hasSort ? [...this.allData[tab] ?? []] : this.allData[tab] ?? [];

    if (hasFilter) {
      result = result.filter((row: SharepointChoiceRow) => {
        // apply all filters
        for (let field in filter) {
          // user hidden columns to skip filters
          if (this.allowHideColumns && this.hiddenColumns[tab]?.includes(field))
            continue;

          let ops = filter[field];
          for (let op in ops) {
            let value = ops[op];
            // get the field value
            var c = this.fieldValue(row, field);
            // apply the operation, all inverted as its easier to think about what is required
            if (op == 'contains') {
              if (!(c?.toString().toLowerCase().includes(value.toString().toLowerCase())))
                return false;
            } else if (op == 'equals') {
              if (!(c?.toString() === value?.toString() || (value == '(blanks)' && (c === null || c === undefined || c === ''))))
                return false;
            } else if (op == 'greater') {
              if (!(c > value))
                return false;
            } else if (op == 'less') {
              if (!(c === null || c < value))
                return false;
            }
          }
        }
        return true;
      });
    }

    if (hasSort)
      result.sort((a: SharepointChoiceRow, b: SharepointChoiceRow) => {
        for (let s of sort) {
          // user hidden columns to skip sorts
          if (this.allowHideColumns && this.hiddenColumns[tab]?.includes(s.field))
            continue;

          let aValue = this.fieldValue(a, s.field);
          let bValue = this.fieldValue(b, s.field);
          let direction = s.direction == 'desc' ? -1 : 1;

          if (aValue === null && bValue === null)
            continue;
          if (aValue === null)
            return 1;
          if (bValue === null)
            return -1;
          if (aValue < bValue)
            return -1 * direction;
          if (aValue > bValue)
            return 1 * direction;
        }

        return 0;
      });

    // Cache the result
    this._rowsCache.set(tab, result);
    return result;
  }

  renderNode(col: SharepointChoiceColumn, value: any, row: SharepointChoiceRow, rowIndex: number, colIndex: number, childIndex?: number): any {
    if (!col.cellRenderer)
      return null;

    // could return issues if cols hidden and many aren't named at all
    const key = `${row['_tracking']}::${col.field ?? col.headerName ?? ((1000 * colIndex) + (childIndex ?? 0)).toString()}`;
    
    const cached = this._nodeCache.get(key);
    if (cached)
      return cached;

    const node = col.cellRenderer(value, row, rowIndex);
    this._nodeCache.set(key, node);
    return node;
  }

  // Helper method to determine cell type (reduces template complexity)
  getCellType(col: SharepointChoiceColumn, value: any): 'renderer' | 'date' | 'number' | 'boolean' | 'default' {
    // quick checks for type
    if (col.cellRenderer) return 'renderer';
    if (col.filter === 'date') return 'date';
    if (col.filter === 'number') return 'number';
    // slower checks based on value
    switch (typeof value)
    {
      case 'number':
        return 'number';
      case 'boolean':
        return 'boolean';
    }
    // check for date object
    return value instanceof Date ? 'date' : 'default';
  }

  // Helper to format boolean values
  formatBoolean(value: any): string {
    return value === true ? '✔' : value === false ? '✘' : '';
  }

  ceil(number: number): number {
    return Math.ceil(number);
  }
}
