import { Component, OnInit, Input, ViewEncapsulation } from '@angular/core';
import pnp from '@pnp/pnpjs';
import { Logger, LogLevel } from "@pnp/logging";
import { PnPLogging, App } from '../../../App';
import { SharepointChoiceUtils } from 'sharepoint-choice';

@Component({
  selector: 'app-create',
  templateUrl: './example.component.html',
  styleUrls: ['./example.component.scss', '../../styles.scss'],  
  encapsulation: ViewEncapsulation.Emulated
})
export class ExampleComponent implements OnInit {
  @Input() context!: string;

  declare form:any[string];
  declare uned:any[string];
  declare versions:any;
  declare spec:any[string];
  declare stage:string;
  declare userId:number;
  declare perm:any[string];

  declare currentPage:number;
  declare itemsPerPage:number;
  declare orderKey:string;
  declare orderDir:boolean;
  declare loading:boolean;

  private util:SharepointChoiceUtils;
  private data:any;
  private list:string;

  constructor() {
    this.util = new SharepointChoiceUtils(this.context);
    this.context = this.util.context;

    this.spec = {'odata.metadata': this.context};
    this.form = {Attachments: {results:[]}};
    
    this.userId = 0;
    this.perm = {};
    this.versions = [];

    this.currentPage = 1;
		this.itemsPerPage = 50;
    this.orderKey = 'EndDate';
    this.orderDir = false;

    this.loading = true;
    this.data = [];
    this.list = 'List';
  }

  ngOnInit(): void {
    pnp.sp.setup({sp:{baseUrl:this.context}});
    Logger.subscribe(new PnPLogging());
    Logger.activeLogLevel = LogLevel.Warning;

    this.util.permissions().then(r => {
      this.userId = r.userId; 
      this.perm = r.perms;
    });

    this.util.fields(this.list).then(r => {
      this.spec = r;
    });
    
    var id = parseInt(this.util.param('aid'));
    this.stage = this.util.param('stage') || (id > 0 ? 'View' : 'New');
    if (id > 0) {
      this.util.data(id, this.list).then(d => {
        this.form = d;
        this.uned = JSON.parse(JSON.stringify(this.form));
      });
      pnp.sp.web.lists.getByTitle(this.list).items.getById(id).update({Audit:'Opened'}).then(() => {
        this.util.history(id, this.list).then(d => {
          this.versions = d
        });
      });
    }

    pnp.sp.web.lists.getByTitle(this.list).items.getAll(5000).then(data => {
      this.data = data;
      this.loading = false;
    })
  }

  dev(o:any):string {
    return JSON.stringify(o);
  }

  isNew(): boolean {
    return this.stage == 'New';
  }

  cantEdit(): boolean {
    return this.stage != 'New' && this.perm['Owners'];
  }

  async onUp(search:string) {
    var results:any[any] = await this.util.msalApi(
      `guid`,
      `permission`,
      'path',
      App.APIRelease || App.Release);
    // ure results
  }

  async save() {
    this.form.Id = await this.util.save(this.form, this.uned, this.list);
    this.close();
  }

  // redirect
  close(): void {
    document.location.href = this.context;
  }

  rows(status: string, display: boolean): any[any] {
    let data = this.data.filter(d => status == null || d.Status == status);

    if (!display)
      return data;

    if (this.orderKey)
      data.sort((a: any, b: any) => {
        if (a[this.orderKey] < b[this.orderKey]) {
          return -1 * (this.orderDir ? -1 : 1);
        } else if (a[this.orderKey] > b[this.orderKey]) {
          return 1 * (this.orderDir ? -1 : 1);
        } else {
          return 0;
        }
      });

    if (this.itemsPerPage > 0)
      return data.slice((this.currentPage -1) * this.itemsPerPage, this.currentPage * this.itemsPerPage);
    
    return data;
  }

  sort (k: string) {
    if (this.orderKey == k)
      this.orderDir = !this.orderDir;
    else {
      this.orderKey = k;
      this.orderDir = false;
    }
  }

  maxItems(status: string): number {
    return this.rows(status, false).length;
  }

  maxPage(status: string, trim: boolean): number {
    var ret = this.maxItems(status) / this.itemsPerPage;
    if (!trim)
      return ret;
    return Math.floor(ret);
  }

  changePage(to: number) {
    this.currentPage = Math.ceil(to);
  }
}
