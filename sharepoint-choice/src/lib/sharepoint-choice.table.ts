import { Component, Input, ErrorHandler, EventEmitter, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SharepointChoiceColumn, SharepointChoiceFilter, SharepointChoiceTabs, SharepointChoiceRow } from './sharepoint-choice.models';

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
  // site unique name used to store column visibility filter sort preferences
  @Input() prefix: string = document.location.href.split('?')[0].split('#')[0];

  // all data passed in keyed by tab name, not using getter/setter to avoid deep copy and memory issues
  @Input() allData?: SharepointChoiceTabs = {};

  // all columns via getter/setter to avoid chance propogation back outwards of hide state etc
  @Input() set allCols(value: SharepointChoiceColumn[]) {
    this._allCols = value || [];
  }
  get allCols(): SharepointChoiceColumn[] {
    return this._allCols;
  }
  private _allCols: SharepointChoiceColumn[] = [];

  // all tabs via getter/setter to allow dynamic getter based on data if no tabs or invalid tabs passed in
  @Input() set allTabs(value: string[]) {
    this._allTabs = value || [];
  }
  get allTabs(): string[] {
    var tabs = Object.keys(this.allData || {}).filter(k => k && k != 'undefined' && k != 'null');
    // if this._allTabs has at least one valid tab then use it else use dynamic tabs from data
    if (this._allTabs.length > 0 || tabs.length == 0) {
      let validTabs = this._allTabs.filter(t => tabs.includes(t));
      if (validTabs.length > 0 || tabs.length == 0)
        return this._allTabs;
    }
    return tabs;
  }
  private _allTabs: string[] = [];

  // selected tab else use stored value or first tab
  @Input() set selectedTab(value: string | undefined) {
    this._selectedTab = value;
  }
  get selectedTab(): string | undefined {
    return this._selectedTab && this.allTabs.includes(this._selectedTab)
      ? this._selectedTab
      : this.getStorage(`Tab`) || this.allTabs[0];
  }
  private _selectedTab?: string;

  // page size via getter/setter to allow storage
  @Input() set pageSize(value: number) {
    this._pageSize = value;
  }
  get pageSize(): number {
    return this._pageSize || this.getStorage(`Size`) || 250;
  }
  private _pageSize?: number;

  // simple variables not needing getter/setter
  @Input() loading: boolean = true;
  @Input() tableHeight: string = 'calc(100vh - 360px)';
  @Input() allowSelection: boolean = false;

  // events
  @Output() selected = new EventEmitter<SharepointChoiceRow[]>();
  @Output() rowClicked = new EventEmitter<{ row: SharepointChoiceRow, target: HTMLElement|EventTarget|null }>();

  pageNumber = 1;
  editColumns = false;

  changeTab(tab: string): void {
    this.selectedTab = tab;
    this.pageNumber = 1;
    this.setStorage(`Tab`, this.selectedTab);
    this.selected.emit([]);
  }

  changeSort(col: SharepointChoiceColumn): void {
    var sort = this.getStorage(`Sort-${this.selectedTab}`) ?? [];
    if (col.sortable === false || !col.field)
      return;
    var asc = sort.indexOf(col.field);
    var desc = sort.indexOf('!' + col.field);
    if (asc == -1 && desc == -1) {
      sort.push(col.field);
    } else if (asc >= 0)
      sort[asc] = '!' + col.field;
    else
      sort = sort.filter((s: string) => s != ('!' + col.field));
    this.setStorage(`Sort-${this.selectedTab}`, sort);
  }

  niceName(col: SharepointChoiceColumn): string {
    var h = col.field?.substring(col.field?.lastIndexOf('.') + 1);
    return col.headerName ? col.headerName : h ? (h.charAt(0).toUpperCase() + h.slice(1)).replace(/([a-z])([A-Z])/g, '$1 $2') : '';
  }

  filterOpts(field?: string): SharepointChoiceFilter {
    var filter = this.selectedTab ? this.getStorage(`Filter-${this.selectedTab}`) ?? {} : {};
    return {
      equals: filter[field ?? '']?.equals ?? null,
      contains: filter[field ?? '']?.contains ?? null,
      greater: filter[field ?? '']?.greater ?? null,
      less: filter[field ?? '']?.less ?? null
    };
  }

  sortContains(field?: string): boolean {
    if (!field || !this.selectedTab)
      return false;
    var sort = this.getStorage(`Sort-${this.selectedTab}`) ?? [];
    return sort.includes(field);
  }

  changeFilter(col: SharepointChoiceColumn, op: string, event: Event): void {
    if (!col.field || !this.selectedTab)
      return;
    var filter = this.getStorage(`Filter-${this.selectedTab}`) ?? {};
    var value: any = null;
    if (event && event.target) {
      value = event.target['value'];
      if (event.target['type'] == 'date' && value)
        value = new Date(value);
    }

    if (!filter[col.field])
      filter[col.field] = {};

    if (value === undefined || value === null || value === '') {
      delete filter[col.field][op];
      if (Object.keys(filter[col.field]).length == 0)
        delete filter[col.field];
    } else
      filter[col.field][op] = value;
    this.setStorage(`Filter-${this.selectedTab}`, filter);
  }

  getStorage(key: string): any {
    var s = localStorage.getItem(`SharepointTable-${this.prefix}`) || '{}';
    return JSON.parse(s)[key];
  }

  setStorage(key: string, value: any) {
    var s = JSON.parse(localStorage.getItem(`SharepointTable-${this.prefix}`) || '{}');
    s[key] = value;
    localStorage.setItem(`SharepointTable-${this.prefix}`, JSON.stringify(s));
  }

  toggleFilter(col: SharepointChoiceColumn, event: Event): void {
    if (!col.field || !this.selectedTab)
      return;
    col._filtervisible = !col._filtervisible;
    event.stopPropagation();
  }

  toggleColumn(col: SharepointChoiceColumn): void {
    if (typeof col.hide == 'function' || !col.field)
      return;

    // because visibility is toggled after click function triggered then we need to reverse the value here
    let hide = !col.hide;

    var cols = this.getStorage(`Hide-${this.selectedTab}`) || [];
    if (hide)
      cols.push(col.field);
    else
      cols = cols.filter((c: string) => c != col.field);
    this.setStorage(`Hide-${this.selectedTab}`, cols);
  }

  flds(tab?: string) {
    // load the column visibility from storage and update allCols, only set hidden dont
    var cols = this.getStorage(`Hide-${this.selectedTab}`) || [];
    var all = this.allCols ?? [];
    all.forEach(c => {
      // top level column def
      if (!c.children) {
        if (cols.includes(c.field ?? '') && typeof c.hide != 'function')
          c.hide = true;
      }
      // load children
      c.children?.forEach((i: SharepointChoiceColumn) => {
        if (cols.includes(i.field ?? '') && typeof i.hide != 'function')
          i.hide = true;
      });
    });
    if (!tab)
      return all;
    return all.filter(c => {
      return !c.hide || (typeof c.hide == 'function' && !c.hide(tab))
    }).map(c => {
      if (c.children) {
        let nc = { ...c };
        nc.children = nc.children?.filter(ch => !ch.hide || (typeof ch.hide == 'function' && !ch.hide(tab)));
        return nc;
      }
      return c;
    }).filter(c => {
      if (c.children)
        return c.children.length > 0;
      return true;
    });
  }

  startResize(event: MouseEvent, col: SharepointChoiceColumn): void {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.pageX;
    const startWidth = col.width || 100;

    const onMouseMove = (e: MouseEvent) => {
      const newWidth = startWidth + (e.pageX - startX);
      col.width = newWidth;
    }
    const onMouseUp = (e: MouseEvent) => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    }
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
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

  distinctContent(field?: string): any[] {
    if (!field || !this.allData || !this.selectedTab || !this.allData[this.selectedTab])
      return [];
    let values: any[] = [];
    this.allData[this.selectedTab].forEach(d => {
      var c = this.fieldValue(d, field);
      if (c !== null && c !== undefined && c !== '' && !values.includes(c))
        values.push(c);
    });
    return values;
  }

  rows(tab?: string): SharepointChoiceRow[] {
    if (!this.allData || !tab || !this.allData[tab])
      return [];
    var sort = this.getStorage(`Sort-${tab}`) ?? [];
    var filter = this.getStorage(`Filter-${tab}`) ?? {};
    return this.allData[tab]
      .filter((row: SharepointChoiceRow) => {
        // apply all filters
        for (let field in filter) {
          let ops = filter[field];
          for (let op in ops) {
            let value = ops[op];
            if (value == null || value == '')
              continue;
            // get the field value
            var c = this.fieldValue(row, field);
            // apply the operation
            if (op == 'contains') {
              if (!c || !c.toString().toLowerCase().includes(value.toString().toLowerCase()))
                return false;
            } else if (op == 'equals') {
              if (c == null || c.toString() != value.toString())
                return false;
            } else if (op == 'greater') {
              if (c == null || c < value)
                return false;
            } else if (op == 'less') {
              if (c == null || c > value)
                return false;
            }
          }
        }
        return true;
      })
      .sort((a: SharepointChoiceRow, b: SharepointChoiceRow) => {
        for (let s of sort) {
          if (!s)
            continue;

          let desc = s.startsWith('!');
          let field = desc ? s.substring(1) : s;
          let aValue = this.fieldValue(a, field);
          let bValue = this.fieldValue(b, field);

          if (aValue === null && bValue === null)
            continue;
          if (aValue === null)
            return 1;
          if (bValue === null)
            return -1;
          if (aValue < bValue)
            return desc ? 1 : -1;
          if (aValue > bValue)
            return desc ? -1 : 1;
        }

        return 0;
      });
  }

  ceil(number: number): number {
    return Math.ceil(number);
  }

  setPageSize(): void {
    this.pageNumber = 1;
    this.setStorage(`Size`, this.pageSize);
  }

  handleCellClick(cellClicked: Function | undefined, row: SharepointChoiceRow, event: Event): void {
    if (cellClicked) {
      cellClicked(row, event.target);
    } else {
      this.rowClicked.emit({ row, target: event.target });
    }
  }

  selectionChanged(row: SharepointChoiceRow): void {
    if (!this.allowSelection || !this.selectedTab)
      return;
    row.selected = !row.selected;
    let selectedRows = this.rows(this.selectedTab).filter(r => r.selected);
    this.selected.emit(selectedRows);
  }
}