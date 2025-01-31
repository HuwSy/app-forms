import { Component, Input, OnInit } from '@angular/core';
import { SharepointChoiceUtils } from 'sharepoint-choice';
import { App, AngularLogging } from '../../../App';

import { ErrorHandler, ElementRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SharepointChoiceComponent } from 'sharepoint-choice';

@Component({
  selector: 'app-hello-world-web-part',
  templateUrl: './hello-world-web-part.component.html',
  styleUrls: ['./hello-world-web-part.component.scss'],
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    SharepointChoiceComponent
  ],
  providers: [{
    provide: ErrorHandler,
    useClass: AngularLogging
  }]
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
  declare tabs:any;
  declare files:any;

  constructor(private elRef: ElementRef) {
    // read attribute as Component bind doesnt trigger @Input
    this.description = this.description || this.elRef.nativeElement.getAttribute('description');
    this.context = this.context || this.elRef.nativeElement.getAttribute('context');
  }

  ngOnInit() {
    this.tabs = [
      {tab: 'New', display: 'Submission', status: 'Draft', owner: 'Visitors'},
      {tab: 'Close', display: 'Close', status: 'Closing', owner: 'Members'},
      {tab: 'Audit', display: 'Completed', status: 'Completed', owner: null}
    ];

    this.files = {Submission: {results:[]}};
    for (var f in this.tabs)
      this.files[this.tabs[f].tab] = {results:[]};

    this._spUtils = new SharepointChoiceUtils(this.context);

    var id = parseInt(this._spUtils.param('aid'));
    this.dashboard = !(id > 0 || id === 0);

    this.list = 'List';
    this.spec = {};

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
    this.form = {Status: 'Draft'};
    this.uned = {};
    this.versions = [];
    this.stage = this._spUtils.param('stage') || (id > 0 ? 'View' : 'New');
    
    if (!this.dashboard) {
      this.moreData();
      if (id > 0) {
        this._spUtils.data(id, this.list).then(async d => {
          this.form = d;
          this.uned = JSON.parse(JSON.stringify(this.form));
      
          this._spUtils.sp.web.lists.getByTitle(this.list).items.getById(id).versions.top(5000)().then(d => {
            this.versions = d
          });

          var f = await this.getFolder();
          if (f != null)
            for (var o in this.files)
              this.files[o].results = await this._spUtils.getFiles(f, o);
        });
      }
    }
  }

  // Dashboard
  // load data
  async loadData(restart: boolean) {
    var cur = JSON.parse(localStorage.getItem(`${this._prefix}-${App.AppName}-${App.Release}-${this._spUtils.context}`) || '{}');
    for(var f in cur) {
      this[f] = cur[f];
    }
  
    if (restart === true) {
      this.loading = true;
      this.currentPage = 1;
    }

    this.data = await this._spUtils.sp.web.lists.getByTitle(this.list).items.filter(``).select("Id", "Created", "Title", "Modified").orderBy("Modified", true).top(5000)();
    
    // data adjusts, for display, searches etc
    this.data.forEach(r => {
      r.title = (r.Title || '').toLowerCase();
    });

    this.loading = false;
  }

  // save specific filter field
  saveFilter(f, r) {
    var cur = JSON.parse(localStorage.getItem(`${this._prefix}-${App.AppName}-${App.Release}-${this._spUtils.context}`) || '{}');
    cur[f] = this[f];
    localStorage.setItem(`${this._prefix}-${App.AppName}-${App.Release}-${this._spUtils.context}`, JSON.stringify(cur));

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
    var results:any[any] = await this._spUtils.callApi('',
      '',
      '',
      '',
      'GET',
      null,
      'json');

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

  async getFolder(needsCreating?:boolean) {
    if (this.form.Storage && this.form.Storage.Url)
      return this.form.Storage.Url;
    
    try {
      let root = await this._spUtils.getRoot('Documents');
      let path = `${root}/${this.form.Id}`;

      if (needsCreating)
        await this._spUtils.ensurePath(path, this._spUtils.context.length < 2 ? 2 : 4);

      return document.location.origin + path;
    } catch (e) {
      alert("Unable to access documents area.");
    }
      
    return null;
  }

  async saveFiles(o:string) {
    if (this.form.Storage.Url == null)
      return;
    
    var req = {Url: `${document.location.href.split('?')[0]}?aid=${this.form.Id}`, Description: `REF ${this.form.Id}`};

    await this._spUtils.ensurePath(this.form.Storage.Url + '/' + o, this._spUtils.context.length < 2 ? 2 : 4);

    await this._spUtils.saveFiles(this.form.Storage.Url, o, req, this.files[o], undefined);
  }

  neededStage(stage:string):boolean {
    if (this.stage == stage)
      return true;
    switch (stage) {
    }
    return true;
  }

  enterKey(e):void {
    if (e.srcElement.tagName != 'TEXTAREA')
      e.preventDefault();
  }

  hasPermission():boolean {
    try {
      switch (this.form.Status) {
        case 'Reject':
        case 'Completed':
          return false;
      }
      // is the owner of the task group
      return this.perm[this.tabs.filter(tab => tab.status == this.form.Status)[0].owner];
    } catch (e) {
      return false;
    }
  }

  // save
  async save(status):Promise<void> {
    if (this.stage == 'View')
      return;

    this.form.Audit = status == 'Unread' ? 'Unread' : status ? 'Completed' : 'Updated';
    if (status == 'Unread')
      status = undefined;

    this.form.Id = await this._spUtils.save(this.form, this.uned, this.list);

    // update versions to abuse its user name processing later
    //this.versions = await pnp.sp.web.lists.getByTitle(this.list).items.getById(this.form.Id).versions.top(5000).get();

    // handle approval of task for next stage
    switch (status) {
      case 'Approved':
        this.form.Rejection = null;
        for (var i = 0; i < this.tabs.length; i++)
          if (this.tabs[i].status == this.form.Status)
            break;
        for (i++; i < this.tabs.length; i++)
          if (this.neededStage(this.tabs[i].tab)) {
            status = this.tabs[i].status;
            break;
          }
        break;
      case 'Reject':
        var reason = prompt("Please provide a rejection reason:");
        if (reason == null || reason == '')
          return;
        this.form.Rejection = `${this.versions[0].Editor.LookupValue} (${(new Date()).toString().split(' GMT')[0]}): ${reason}`;
        for (var i = 0; i < this.tabs.length; i++)
          if (this.tabs[i].status == this.form.Status)
            break;
        for (i--; i >= 0; i--)
          if (this.neededStage(this.tabs[i].tab)) {
            status = this.tabs[i].status;
            break;
          }
        break;
    }
    
    // if no folder path calculate and save to request
    if (!this.form.Storage || !this.form.Storage.Url) {
      // may fail to save on long paths, continue anyway
      this.form.Storage = {Url: await this.getFolder(true), Description: 'here'};
      if (this.form.Storage.Url != null)
        await this._spUtils.sp.web.lists.getByTitle(this.list).items.getById(this.form.Id).update({ Storage: this.form.Storage });
    }

    // save relevant files
    for (var o in this.files)
      if (this.files[o].results.length > 0)
        await this.saveFiles(o);

    await this._spUtils.save(this.form, this.uned, this.list);

    if (this.hasPermission() && this.form.Rejection == null) {
      document.location.href = `${document.location.href.split('?')[0]}?aid=${this.form.Id}`;
      return;
    }

    this.close();
  }

  // close
  close(): void {
    document.location.href = document.location.href.split('?')[0];
  }
}
