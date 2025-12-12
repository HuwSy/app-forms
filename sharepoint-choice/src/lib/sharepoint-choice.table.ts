import { Component, Input, ErrorHandler, EventEmitter, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SharepointChoiceColumn, SharepointChoiceFilter, SharepointChoiceSort, SharepointChoiceHide, SharepointChoiceTabs, SharepointChoiceRow } from './sharepoint-choice.models';

@Component({
  selector: 'app-table',
  templateUrl: './sharepoint-choice.table.html',
  styleUrls: ['../styles.scss'],
  standalone: true,
  imports: [
    CommonModule,
    FormsModule
  ],
  providers: [{
    provide: ErrorHandler
  }]
})
export class SharepointChoiceTable {
  // all data passed in keyed by tab name (not via signals which reduce performance on large data sets)
  @Input() set allData(value: SharepointChoiceTabs) {
    this._allData = value;
    // Clear cache when data changes
    this._rowsCache.clear();
    // Mark as loaded after first data set and at least one tab with data
    if (value && Object.keys(value).length > 0 && Object.values(value).some(tabData => tabData && tabData.length > 0)) {
      this.hasLoadedData = true;
    }
  }
  get allData(): SharepointChoiceTabs {
    return this._allData || {};
  }
  private _allData: SharepointChoiceTabs = {};

  // all columns
  @Input() set allCols(value: SharepointChoiceColumn[]) {
    this._allCols = value;
  }
  get allCols(): SharepointChoiceColumn[] {
    return this._allCols || [];
  }
  private _allCols: SharepointChoiceColumn[] = [];

  // all tabs or derive from all data
  @Input() set allTabs(value: string[]) {
    this._allTabs = value;
  }
  get allTabs(): string[] {
    var tabs = Object.keys(this.allData).filter(k => k && k != 'undefined' && k != 'null');
    if (!this._allTabs || this._allTabs.length == 0)
      return tabs;
    if (tabs.length == 0)
      return this._allTabs;
    if (this._allTabs.filter(t => tabs.includes(t)).length > 0)
      return this._allTabs;
    return tabs;
  }
  private _allTabs: string[] = [];

  // selected tab else use stored value or first tab
  @Input() set selectedTab(value: string | undefined) {
    this._selectedTab = value;
    this.setStorage(`Tab`, this.selectedTab);
  }
  get selectedTab(): string | undefined {
    var tab = this._selectedTab || this.getStorage(`Tab`);
    if (this.allTabs.length == 0)
      return tab;
    if (tab && this.allTabs.includes(tab))
      return tab;
    return this.allTabs[0];
  }
  private _selectedTab?: string;

  // page size via getter/setter to allow storage
  @Input() set pageSize(value: number) {
    this._pageSize = value;
    this.setStorage(`Size`, this.pageSize);
  }
  get pageSize(): number {
    return this._pageSize || this.getStorage(`Size`) || 250;
  }
  private _pageSize?: number;

  // simple inputs that dont need getter/setter
  @Input() prefix: string = document.location.href.split('?')[0].split('#')[0];
  @Input() loading: boolean = false;
  @Input() tableHeight: string = 'calc(100vh - 360px)';
  @Input() allowSelection: boolean = false;

  // outbound events or pseudo callbacks for await support
  @Input() rowClicking: Function = async (row: SharepointChoiceRow, target: HTMLElement|EventTarget|null) => {};
  @Output() selected = new EventEmitter<{ data: SharepointChoiceRow[], tab: string }>();
  @Output() rowClicked = new EventEmitter<{ row: SharepointChoiceRow, target: HTMLElement|EventTarget|null }>();

  // internal state
  pageNumber = 1;
  editColumns = false;
  hasLoadedData = false;

  // Memoization cache for filtered/sorted rows
  private _rowsCache: Map<string, SharepointChoiceRow[]> = new Map();

  // stored sort/filter/hidden columns via getter/setter with storage and recall
  set sort(value: SharepointChoiceSort) {
    this._sort = value;
    this.setStorage(`Sort`, this.sort);
    // Clear cache when sort changes
    if (this.selectedTab) {
      this._rowsCache.delete(this.selectedTab);
    }
  }
  get sort(): SharepointChoiceSort {
    return this._sort || this.getStorage(`Sort`) || {};
  }
  private _sort: SharepointChoiceSort = {};

  set filter(value: SharepointChoiceFilter) {
    this._filter = value;
    this.setStorage(`Filter`, this.filter);
    // Clear cache when filter changes
    if (this.selectedTab) {
      this._rowsCache.delete(this.selectedTab);
    }
  }
  get filter(): SharepointChoiceFilter {
    return this._filter || this.getStorage(`Filter`) || {};
  }
  private _filter: SharepointChoiceFilter = {};

  set hiddenColumns(value: SharepointChoiceHide) {
    this._hiddenColumns = value;
    this.setStorage(`Hide`, this.hiddenColumns);
  }
  get hiddenColumns(): SharepointChoiceHide {
    return this._hiddenColumns || this.getStorage(`Hide`) || {};
  }
  private _hiddenColumns: SharepointChoiceHide = {};

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
    this.pageNumber = 1;
    this.selected.emit({ data: [], tab: this.selectedTab });
  }

  // selection change toggles row selected state and emits selected rows
  selectionChanged(row: SharepointChoiceRow): void {
    if (!this.selectedTab)
      return this.selected.emit({ data: [], tab: this.selectedTab });
    row.selected = !row.selected;
    // if there is filtering on selected, clear cache to refresh rows
    if (!!(this.filter[this.selectedTab]?.['selected']))
      this._rowsCache.delete(this.selectedTab);
    let selectedRows = this.rows(this.selectedTab).filter(r => r.selected);
    this.selected.emit({ data: selectedRows, tab: this.selectedTab });
  }

  // sorts, filters, column visibility changes
  sortChange(col: SharepointChoiceColumn): void {
    if (!col.field || col.sortable === false || !this.selectedTab)
      return;

    // Get current sort state
    const currentSort = { ...this.sort };
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

  filterChange(col: SharepointChoiceColumn, op: string, event: Event): void {
    if (!col.field || col.filter == 'none' || !this.selectedTab)
      return;

    var value: any = null;
    if (event && event.target) {
      value = event.target['value'];
      if (event.target['type'] == 'date' && value && !value.toJSON)
        value = new Date(value);
    }

    // Get current filter state
    const currentFilter = { ...this.filter };
    if (!currentFilter[this.selectedTab])
      currentFilter[this.selectedTab] = [];

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
  }

  columnToggle(col: SharepointChoiceColumn, tab: string): void {
    if (!col.field)
      return;

    let curr = this.isHidden(col, tab);
    const currentHidden = { ...this.hiddenColumns };
    
    if (!currentHidden[tab])
      currentHidden[tab] = [];
    
    if (!curr)
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
    }
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }

  // handle cell click or row clicks
  async handleCellClick(cellClicked: Function | undefined, row: SharepointChoiceRow, event: Event): Promise<void> {
    if (cellClicked) {
      let c = cellClicked(row, event.target);
      if (c instanceof Promise)
        await c;
    } else {
      if (this.rowClicked) {
        // unfortunately emit doesnt support await
        this.rowClicked.emit({ row, target: event.target });
      }
      if (this.rowClicking) {
        let c = this.rowClicking(row, event.target);
        if (c instanceof Promise)
          await c;
      }
    }
    // Clear cache in case callback modified row data that affects filtering/sorting
    if (this.selectedTab)
      this._rowsCache.delete(this.selectedTab);
  }

  isHidden(col: SharepointChoiceColumn, tab: string): boolean {
    if (typeof col.hide == 'function' && col.hide(tab))
      return true;
    if (!col.field || !this.hiddenColumns[tab])
      return false;
    return this.hiddenColumns[tab].includes(col.field);
  }

  // get columns based on tab and hidden state
  fields(tab: string): SharepointChoiceColumn[] {
    let cols = this.hiddenColumns[tab] || [];
    // hide based on tab function or hide state
    return this.allCols.filter(c => {
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
  }

  // utility functions
  niceName(col: SharepointChoiceColumn): string {
    if (col.headerName)
      return col.headerName;
    var h = col.field?.substring(col.field?.lastIndexOf('.') + 1);
    if (h)
      col.headerName = h.charAt(0).toUpperCase() + h.slice(1).replace(/([a-z])([A-Z])/g, '$1 $2');
    return col.headerName || '';
  }

  fieldValue(row: SharepointChoiceRow, field?: string): any {
    if (!field)
      return null;
    try {
      var f = field.split('.');
      var c = row[f[0]];
      for (let i = 1; i < f.length; i++)
        c = c ? c[f[i]] : null;
      return c;
    } catch (e) {
      return null;
    }
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

  // get rows based on tab, sort and filter
  rows(tab: string): SharepointChoiceRow[] {
    if (!this.allData || !tab || !this.allData[tab])
      return [];

    // Return cached result if available (cache cleared on filter/sort changes)
    const cached = this._rowsCache.get(tab);
    if (cached)
      return cached;

    var filter = this.filter[tab] || {};
    var sort = this.sort[tab] || [];

    const result = this.allData[tab]
      .filter((row: SharepointChoiceRow) => {
        // apply all filters
        for (let field in filter) {
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
              if (!(c?.toString() === value.toString() || (value == '(blanks)' && (c === null || c === undefined || c === ''))))
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
      })
      .sort((a: SharepointChoiceRow, b: SharepointChoiceRow) => {
        for (let s of sort) {
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

  // Helper method to determine cell type (reduces template complexity)
  getCellType(col: SharepointChoiceColumn, value: any): 'renderer' | 'date' | 'number' | 'boolean' | 'default' {
    if (col.cellRenderer) return 'renderer';
    if (col.filter === 'date' || value?.toJSON) return 'date';
    if (col.filter === 'number' || typeof value === 'number') return 'number';
    if (typeof value === 'boolean') return 'boolean';
    return 'default';
  }

  // Helper to format boolean values
  formatBoolean(value: any): string {
    return value === true ? '✔' : value === false ? '✘' : '';
  }

  ceil(number: number): number {
    return Math.ceil(number);
  }

}
