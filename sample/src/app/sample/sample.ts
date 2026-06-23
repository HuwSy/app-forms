import { Component, Input, OnInit } from '@angular/core';
import { SharepointChoiceUtils, SharepointChoiceLogging } from 'sharepoint-choice';

import { ElementRef, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SharepointChoiceComponent, SharepointChoiceTable } from 'sharepoint-choice';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-groups";

@Component({
  selector: 'app-sample',
  templateUrl: './sample.html',
  styleUrls: ['./sample.scss'],
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    SharepointChoiceComponent,
    SharepointChoiceTable
  ]
})
export class SampleComponent implements OnInit {
  @Input() description!: string;
  @Input() context!: string;

  // Common
  dashboard:boolean = true;
  userId:number = 0;
  perm:any[string] = {};
  list:string = 'List';
  spec:any[string] = {};
  
  private _spUtils: SharepointChoiceUtils;
  private _log: SharepointChoiceLogging;

  // Dashboard
  loading:boolean = true;
  data:any = [];
  search:any = {};
  allEditing = false;

  // Form
  form:any[string] = {Status: 'Draft'};
  versions:any = [];
  uned:any[string] = {};
  stage:string = 'View';
  tabs = [
    {tab: 'New', display: 'Submission', status: 'Draft', owner: 'Visitors'},
    {tab: 'Close', display: 'Close', status: 'Closing', owner: 'Members'},
    {tab: 'Audit', display: 'Completed', status: 'Completed', owner: null}
  ];
  files:any = {Submission: {results:[]}};

  titleSearch = {
    parent: this,
    search: async (text: string) => {
      // return matching rows for inline search
      return [
        { Title: 'Alpha' },
        { Title: 'Beta' },
        { Title: 'Gamma' }
      ].filter(x => x.Title.toLowerCase().includes(text.toLowerCase()));
    },
    select: async (row: any) => {
      console.log('selected row', row);
    }
  };

  fileOptions = {
    extract: true,
    check: true,
    accept: '.pdf,.docx,.xlsx',
    download: true,
    uploadonly: false,
    archive: 'Archived',
    view: 0,
    doctypes: ['Invoice', 'Contract'],
    doctype: 'DocumentType',
    notes: 'Notes',
    spec: {
      DocumentType: {
        TypeAsString: 'Choice',
        InternalName: 'DocumentType',
        Title: 'Document Type',
        Choices: ['Invoice', 'Contract', 'Receipt']
      },
      Notes: {
        TypeAsString: 'Text',
        InternalName: 'Notes',
        Title: 'Notes'
      },
      Archived: {
        TypeAsString: 'Boolean',
        InternalName: 'Archived',
        Title: 'Archived'
      }
    }
  };
  
  constructor(private elRef: ElementRef, private chRef: ChangeDetectorRef) {
    // read attribute as Component bind doesnt trigger @Input
    this.description = this.description || this.elRef.nativeElement.getAttribute('description');
    this.context = this.context || this.elRef.nativeElement.getAttribute('context');
    this._spUtils = new SharepointChoiceUtils(this.context);
    this._log = new SharepointChoiceLogging();
  }

  async ngOnInit() {
    try {
      for (var f in this.tabs)
        this.files[this.tabs[f].tab] = {results:[]};
  
      var id = parseInt(this._spUtils.param('aid') || '0');
      this.dashboard = !(id > 0 || id === 0);
      
      this.stage = (id > 0 ? 'View' : 'New');
  
      let p = await this._spUtils.permissions();
      this.userId = p.userId; 
      this.perm = p.perms;
      this.chRef.detectChanges();
      
      let s = await this._spUtils.fields(this.list);
      this.spec = s;
      this.chRef.detectChanges();
      
      if (this.dashboard)
        this.loadData(false);
      else
        this.loadForm(id);
    } catch (e) {
      this._log.handleError(e);
    }
  }

  loadForm(id) {
    this.moreData();
    if (id <= 0)
      return;
    this._spUtils.data(id, this.list).then(async d => {
      this.form = d;
      this.uned = JSON.parse(JSON.stringify(this.form));
      this.loading = false;
      this.chRef.detectChanges();
        
      this._spUtils.version(id, this.list).then(d => {
        this.versions = d;
        this.chRef.detectChanges();
      });
  
      var f = await this.getFolder();
      if (f != null)
        for (var o in this.files)
          this.files[o].results = await this._spUtils.getFiles(f, o);
        this.chRef.detectChanges();
    });
  }

  async loadData(restart: boolean) {
    this.data = await this._spUtils.sp.web.lists.getByTitle(this.list).items.filter(``).select("Id", "Created", "Title", "Modified").orderBy("Modified", true).top(5000)();
    this.loading = false;
    this.chRef.detectChanges();
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
    
    this.chRef.detectChanges();
  }

  lookupSearch = {
    search: async (query: string) => {
      const r = await this._spUtils.search(query, 20);
      return r.PrimarySearchResults.map((x:any) => ({
        Id: x.Id,
        Title: x.Title
      }));
    },
    select: (item: any) => {
      this.form.LookupField = item;
    }
  };
  
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
    
    let root = await this._spUtils.getRoot('Documents');
    let path = `${root}/${this.form.Id}`;

    if (needsCreating)
      await this._spUtils.ensurePath(path, this._spUtils.context.length < 2 ? 2 : 4);

    return document.location.origin + path;
  }

  async saveFiles(o:string) {
    if (this.form.Storage.Url == null)
      return;
    
    var req = {Url: `${document.location.href.split('?')[0]}?aid=${this.form.Id}`, Description: `REF ${this.form.Id}`};

    await this._spUtils.ensurePath(this.form.Storage.Url + '/' + o, this._spUtils.context.length < 2 ? 2 : 4);

    await this._spUtils.saveFiles(this.form.Storage.Url, o, req, this.files[o], undefined);
    
    this.chRef.detectChanges();
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

  hyperlink(rowData) {
    return "?aid="+rowData.Id;
  }
  
  rowClicked = async (rowData: any, target: HTMLElement | EventTarget | undefined) => {
    console.log('rowClicked callback', rowData, target);
    return true;
  };
  
  onSelected(event: { data: any[]; tab: string | undefined }) {
    console.log('selected output', event);
  }

  onCleared() {
    this.search.Title = undefined;
  }

  onClicked(event: { row: any; target: HTMLElement | EventTarget | undefined }) {
    console.log('clicked output', event);
  }

  hasPermission():boolean {
    try {
      // some status have no matching tab stages
      switch (this.form.Status) {
        case 'Reject':
        case 'Completed':
          return false;
      }
      // is the owner of the task group
      return this.perm[this.tabs.filter(tab => tab.tab == this.stage)[0]?.owner ?? ''];
    } catch (e) {
      return false;
    }
  }

  // save
  async save(status):Promise<void> {
    try {
      if (this.stage == 'View')
        return;
  
      this.form.Audit = status == 'Unread' ? 'Unread' : status ? 'Completed' : 'Updated';
      if (status == 'Unread')
        status = undefined;
  
      this.form.Id = await this._spUtils.save(this.form, this.uned, this.list);
      
      this.chRef.detectChanges();
  
      this.versions = await this._spUtils.sp.web.lists.getByTitle(this.list).items.getById(this.form.Id).versions.top(5000).get();
  
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
  
      this.chRef.detectChanges();
      this.close();
    } catch (e) {
      this._log.handleError({
        error: e,
        form: this.form,
        files: this.files
      });
    }
  }

  // close
  close(): void {
    document.location.href = document.location.href.split('?')[0];
  }
}
