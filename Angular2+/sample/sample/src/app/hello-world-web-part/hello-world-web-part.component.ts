import { Component, Input, OnInit, ViewEncapsulation } from '@angular/core';
import pnp from '@pnp/pnpjs';
import { Logger, LogLevel } from "@pnp/logging";
import { PnPLogging } from '../util/PnPLogging';
import { SharepointChoiceUtils } from 'sharepoint-choice';
import { App } from '../util/App';

@Component({
  selector: 'app-hello-world-web-part',
  templateUrl: './hello-world-web-part.component.html',
  styleUrls: ['./hello-world-web-part.component.scss'],
  encapsulation: ViewEncapsulation.Emulated
})
export class HelloWorldWebPartComponent implements OnInit {
  @Input() description!: string;
  @Input() context!: string;

  // Common
  declare dashboard:boolean;
  declare userId:number;
  declare perm:any[string];
  declare list:string;
  declare spec:any[string];
  private _spUtils: SharepointChoiceUtils;

  // Dashboard
  declare searchText:string;
  declare currentPage:number;
  declare itemsPerPage:number;
  declare orderKey:string;
  declare orderDir:boolean;
  declare refresh:number;
  declare loading:boolean;
  declare status:any[string];
  declare data:any;
  declare selected:string;
  private _prefix: string = 'Choice-Filter';

  // Form
  declare form:any[string];
  declare versions:any;
  declare uned:any[string];
  declare stage:string;

  constructor() { }

  ngOnInit() {
    var id = parseInt(this._spUtils.param('aid'));
    this.dashboard = !(id > 0 || id === 0);

    // Common
    this._spUtils = new SharepointChoiceUtils(this.context);

    pnp.sp.setup({sp:{baseUrl:this.context}});
    Logger.subscribe(new PnPLogging());
    Logger.activeLogLevel = LogLevel.Warning;

    this.list = 'List';
    this.spec = {'odata.metadata': this.context};

    this.userId = 0;
    this.perm = {};

    this._spUtils.permissions().then(r => {
      this.userId = r.userId; 
      this.perm = r.perms;
    });
    
    this._spUtils.fields(this.list).then(r => {
      this.spec = r;
      this.status = this.spec['Status']?.Choices;
    });
    
    // Dashboard
    this.selected = '';

    this.searchText = '';
    this.currentPage = 1;
		this.itemsPerPage = 25;
    this.orderKey = null;
    this.orderDir = false;
    this.refresh = -1;
    this.status = [];

    this.loading = true;
    this.data = [];
    if (this.dashboard) {
      this.loadData(false);
      if (this.refresh > 0)
        setInterval(this.loadData, this.refresh * 1000);
    }

    // Form
    this.form = {__metadata:{type:`SP.Data.${this.list}ListItem`}};
    this.uned = {};
    this.versions = [];
    this.stage = this._spUtils.param('stage') || (id > 0 ? 'View' : 'New');
    
    if (!this.dashboard) {
      this.moreData();
      if (id > 0) {
        this._spUtils.data(id, this.list).then(d => {
          this.form = d;
          this.uned = JSON.parse(JSON.stringify(this.form));
        });
        this._spUtils.history(id, this.list).then(d => {
          this.versions = d
        });
      }
    }
  }

  // Dashboard
  // load data
  async loadData(restart: boolean) {
    var cur = JSON.parse(localStorage.getItem(`${this._prefix}-${App.AppName}-${App.Release}-${this.context}`) || '{}');
    for(var f in cur) {
      this[f] = cur[f];
    }
  
    if (restart === true) {
      this.loading = true;
      this.currentPage = 1;
    }

    this.data = await pnp.sp.web.lists.getByTitle(this.list).items.filter(``).select("Id", "Created", "Title", "Modified").orderBy("Modified", true).getAll(5000);
    
    // data adjusts, for display, searches etc
    this.data.forEach(r => {
      r.title = (r.Title || '').toLowerCase();
    });

    this.loading = false;
  }

  // save specific filter field
  saveFilter(f, r) {
    var cur = JSON.parse(localStorage.getItem(`${this._prefix}-${App.AppName}-${App.Release}-${this.context}`) || '{}');
    cur[f] = this[f];
    localStorage.setItem(`${this._prefix}-${App.AppName}-${App.Release}-${this.context}`, JSON.stringify(cur));

    if (r) {
      this.selected = null;
      this.loadData(true);
    }
  }

  // select sub heading
  select(s) {
    if (this.selected == s)
      this.selected = null;
    else
      this.selected = s;
  }

  // filter to show rows
  rows(status: string, display: boolean): any[any] {
    var ths = this;
    var ret = this.data.filter(r => {
      if (status
        && r.Status != status)
        return false;
      if ((ths.searchText || '') != ''
        && !~r.title.indexOf(ths.searchText.toLowerCase()))
        return false;
      return true;
    });
    
    if (!display)
      return ret;

    if (this.orderKey)
      ret.sort((a: any, b: any) => {
        if (a[this.orderKey] < b[this.orderKey]) {
          return -1 * (this.orderDir ? -1 : 1);
        } else if (a[this.orderKey] > b[this.orderKey]) {
          return 1 * (this.orderDir ? -1 : 1);
        } else {
          return 0;
        }
      });

    if (this.itemsPerPage > 0)
      return ret.slice((this.currentPage -1) * this.itemsPerPage, this.currentPage * this.itemsPerPage);
    
    return ret;
  }

  // on clicking headings
  sort (k: string) {
    if (this.orderKey == k)
      this.orderDir = !this.orderDir;
    else {
      this.orderKey = k;
      this.orderDir = false;
    }
    this.saveFilter('orderKey', false);
    this.saveFilter('orderDir', false);
  }

  // last page number
  maxPage(status: string, trim: boolean): number {
    var ret = this.rows(status, false).length / this.itemsPerPage;
    if (!trim)
      return ret;
    return Math.floor(ret);
  }

  // change page
  changePage(to: number) {
    this.currentPage = Math.ceil(to);
  }

  // Form
  // file types for attachments
  prefix():string {
    var fileTypes = ["One", "Two", "Three"];
    return `{"Prefix":${JSON.stringify(fileTypes)}, "TypeAsString": "Attachments"}`; //"TypeAsString": "Attachments" not required except under ng serve
  }

  // additional choice data via api
  async moreData(): Promise<any> {
    var results:any[any] = await this._spUtils.msalApi(App.AzureApp,
      '/Read',
      `MoreData`,
      App.Release);

    results.forEach(d => {
      this.spec.Choices.push(d);
    });
  }
  
  // add subtract repeating sections
  add(f):void {
    if (!this.form[f])
      this.form[f] = [];
    this.form[f].push({});
  }
  sub(f, i:number):void {
    this.form[f].splice(i,1);
    if (this.form[f].length == 0)
      this.form[f] = null;
  }

  // override required
  required(x:string): string {
    var r = JSON.stringify({Required: true});
    if (x == 'Title' && this.stage != 'New')
      return r;
    return '';
  }

  // save
  async save(status):Promise<void> {
    // any additional logic
    this.form.Id = await this._spUtils.save(this.form, this.uned, this.list);
    this.close();
  }

  // close
  close(): void {
    document.location.href = document.location.href.split('?')[0];
  }
}
