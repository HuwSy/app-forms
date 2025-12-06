import { Component, Input, ErrorHandler } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

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
  // all tabs in desired order passed in, if only one then no tabs shown
  @Input() allTabs: string[] = [];
  // default to first tab if not passed in
  @Input() selectedTab: string = this.allTabs.length > 0 ? this.allTabs[0] : '';

  @Input() loading: boolean = true;
  @Input() pageSize: number = 100;
  // all data passed in, keyed by tab name
  @Input() allData: any = {};
  @Input() allCols: any[] = [];
  // site unique name used to store column visibility preferences
  @Input() siteName: string = 'https://default.site/sites/default';

  @Input() rowClicked: Function = (row: any, event: any) => {};
  @Input() tableHeight: string = 'calc(100vh - 330px)';

  /*
    <app-table [allTabs]="tabs"
               [selectedTab]="tabs[0]"
               [allData]="allData"
               [allCols]="allCols"
               [siteName]="context"
               [tableHeight]="'calc(100vh - 330px)'">
    </app-table>
  */
  /* tabs
    [ 'tab1', 'tab2', ... ]
  */
  /* data
    { 
      tab1: [ { title: 'item1', ... }, { title: 'item2', ... } ],
      tab2: [ { title: 'itemA', ... }, { title: 'itemB', ... } ]
    }
  */
  /* cols
    {
      headerName: 'Col title',
      field: 'title',
      headerTooltip: 'some tooltip',
      nowrap: true,
      cellClicked: (value, row, event) => {
        row.selected = !row.selected;
      },
      // must be string template not HTMLElement
      cellRenderer: (val, row, index) => {
        if (row.data == 1)
          return val;

        return `
          ${
            row.selected
            ? '<span>✔</span>'
            : '<span style="font-size: 20px;">☐</span>'
          }
          ${val}
        `;
      },
      filter: 'text'
    }
  */

  pageNumber = 1;
  editColumns = false;
  sort:any = [];
  filter:any = {};

  changeTab(tab: string): void {
    this.selectedTab = tab;
    this.pageNumber = 1;
  }

  changeSort(col) {
    if (col.sortable === false || !col.field)
      return;
    var asc = this.sort.indexOf(col.field);
    var desc = this.sort.indexOf('!' + col.field);
    if (asc == -1 && desc == -1) {
      if (this.sort.length >= 3)
        this.sort.splice(0, 1);
      this.sort.push(col.field);
    } else if (asc >= 0)
      this.sort[asc] = '!' + col.field;
    else
      this.sort = this.sort.filter(s => s != ('!' + col.field));
  }

  niceName(col): string {
    var h = col.field?.substring(col.field?.lastIndexOf('.') + 1);
    return col.headerName ? col.headerName : (h?.charAt(0).toUpperCase() + h?.slice(1)).replace(/([a-z])([A-Z])/g, '$1 $2');
  }

  sortContains(field: string): boolean {
    return this.sort.includes(field);
  }

  changeFilter(col, op: string, event: Event): void {
    var value:any = null;
    if (event && event.target) {
      value = event.target['value'];
      if (event.target['type'] == 'date' && value)
        value = new Date(value);
    }

    if (!this.filter[col.field])
      this.filter[col.field] = {};
    
    if (value === undefined || value === null || value === '') {
      delete this.filter[col.field][op];
      if (Object.keys(this.filter[col.field]).length == 0)
        delete this.filter[col.field];
    } else
      this.filter[col.field][op] = value;
  }

  toggleColumn(col: any): void {
    if (typeof col.hide == 'function' || !col.field)
      return;

    // because visibility is toggled after click function triggered then we need to reverse the value here
    let hide = !col.hide;

    var cols = (localStorage.getItem(`Choice-Filter-${this.siteName}-${this.selectedTab}`) || '').split(',').filter(c => c);
    if (hide)
      cols.push(col.field);
    else
      cols = cols.filter(c => c != col.field);
    localStorage.setItem(`Choice-Filter-${this.siteName}-${this.selectedTab}`, cols.join(','));
  }

  flds(tab?: string) {
    // load the column visibility from storage and update allCols, only set hidden dont
    var cols = (localStorage.getItem(`Choice-Filter-${this.siteName}-${this.selectedTab}`) || '').split(',').filter(c => c);
    var all = this.allCols ?? [];
    all.forEach(c => {
      // top level column def
      if (!c.children) {
        if (cols.includes(c.field ?? '') && typeof c.hide != 'function')
          c.hide = true;
      }
      // load children
      c.children?.forEach((i: any) => {
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
        nc.children = nc.children.filter(ch => !ch.hide || (typeof ch.hide == 'function' && !ch.hide(tab)));
        return nc;
      }
      return c;
    }).filter(c => {
      if (c.children)
        return c.children.length > 0;
      return true;
    });
  }

  startResize(event: MouseEvent, col: any): void {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.pageX;
    const startWidth = col.width || col.renderwidth || 100;

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

  fieldValue(row: any, field: string): any {
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

  distinctContent(field: string): any[] {
    let values: any[] = [];
    (this.allData[this.selectedTab] ?? []).forEach(d => {
      var c = this.fieldValue(d, field);
      if (c != null && !values.includes(c))
        values.push(c);
    });
    return values;
  }

  rows(tab: string): any[] {
    return (this.allData[tab] ?? [])
      .filter((row: any) => {
        // apply all filters
        for (let field in this.filter) { 
          let ops = this.filter[field];
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
      .sort((a: any, b: any) => {
        let s = this.sort[2];
        if (!s)
          return 0;
        let desc = s.startsWith('!');
        let field = desc ? s.substring(1) : s;
        let f = field.split('.');
        let aValue = a[f[0]];
        let bValue = b[f[0]];
        for (let i = 1; i < f.length; i++) {
          aValue = aValue ? aValue[f[i]] : null;
          bValue = bValue ? bValue[f[i]] : null;
        }
        if (aValue < bValue)
          return desc ? 1 : -1;
        if (aValue > bValue)
          return desc ? -1 : 1;
        return 0;
      })
      .sort((a: any, b: any) => {
        let s = this.sort[1];
        if (!s)
          return 0;
        let desc = s.startsWith('!');
        let field = desc ? s.substring(1) : s;
        let f = field.split('.');
        let aValue = a[f[0]];
        let bValue = b[f[0]];
        for (let i = 1; i < f.length; i++) {
          aValue = aValue ? aValue[f[i]] : null;
          bValue = bValue ? bValue[f[i]] : null;
        }
        if (aValue < bValue)
          return desc ? 1 : -1;
        if (aValue > bValue)
          return desc ? -1 : 1;
        return 0;
      })
      .sort((a: any, b: any) => {
        let s = this.sort[0];
        if (!s)
          return 0;
        let desc = s.startsWith('!');
        let field = desc ? s.substring(1) : s;
        let f = field.split('.');
        let aValue = a[f[0]];
        let bValue = b[f[0]];
        for (let i = 1; i < f.length; i++) {
          aValue = aValue ? aValue[f[i]] : null;
          bValue = bValue ? bValue[f[i]] : null;
        }
        if (aValue < bValue)
          return desc ? 1 : -1;
        if (aValue > bValue)
          return desc ? -1 : 1;
        return 0;
      });
  }

  ceil(number: number): number {
    return Math.ceil(number);
  }

}

