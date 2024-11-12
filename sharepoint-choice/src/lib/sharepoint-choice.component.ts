import { Component, OnInit, OnDestroy, Input, ElementRef, ViewEncapsulation } from '@angular/core';
import { UserQuery, User } from "./Models";
import "@pnp/sp/webs";
import { Logger, LogLevel } from "@pnp/logging";
import { Editor, Toolbar } from 'ngx-editor';
import * as MsgReader from '@sharpenednoodles/msg.reader-ts';
import { readEml } from 'eml-parse-js';
import * as JSZip from 'jszip';
import { Subject } from 'rxjs';
import { debounceTime, distinctUntilChanged } from 'rxjs/operators';
import { SharepointChoiceUtils } from './sharepoint-choice.utils';
import { App } from './App';

@Component({
  selector: 'app-choice',
  templateUrl: './sharepoint-choice.component.html',
  styleUrls: ['../styles.scss'],
  encapsulation: ViewEncapsulation.Emulated
})
export class SharepointChoiceComponent implements OnInit, OnDestroy {
  @Input() form!: Array<string>; // form containing field
  @Input() field!: string; // internal field name on form object, used for push back and against spec

  @Input() spec!: Array<Array<string>>; // spec of field loaded from list
  @Input() override!: string; // manually override any spec above. sent as string as passing object kills large form performance

  @Input() disabled!: boolean; // get disabled state from outside
  @Input() onchange!: Function; // onchange trigger a function(this)

  @Input() text!: { // override text for field
    pattern?:string, // regex pattern for validation

    search?:Function, // search via api for drop down options
    select?:Function, // upon selection in drop down call back function
    parent?:any // parent object that the control belongs to for call backs
  };

  @Input() select!: { // override select for field
    none?: string, // none option text instead of null
    other?: string, // Other fill-in option text, will override to allow other

    filter?: Function // filter choices by a function
    parent?:any // parent object that the control belongs to for call backs
  };
  
  @Input() file!: { // override file for field
    extract?: boolean, // extract files from zip and email
    primary?: string, // primary field name
    doctypes?: Array<string>, // document types
    doctype?: string, // document type field name
    notes?: string, // notes field name
    archive?: string, // archive field name
    view?: number, // view type
    accept?: string, // accept file types attribute
    download?: boolean, // force download of files
    uploadonly?: boolean // only upload files
  };

  declare editor: Editor;
  declare toolbar: Toolbar;
  declare tooltip: boolean;
  declare filesOver: boolean;
  declare name: string;
  declare users: Array<User>;
  declare display: any;
  declare loading: Array<number>;
  declare UserQuery: UserQuery;
  declare filterMulti: string;
  declare unused: string;
  declare results: any[any];
  declare pos: number;

  public textKey: Subject<string> = new Subject<string>();
  public userKey: Subject<string> = new Subject<string>();

  constructor(
    private elRef: ElementRef
  ) {
    if (!this.text)
      this.text = { };
    if (!this.select)
      this.select = { };
    if (!this.file)
      this.file = { };

    // rich text field
    this.editor = new Editor();
    // rtf menu items
    this.toolbar = [
      ['text_color', 'background_color'],
      ['bold', 'italic', 'underline', 'strike'],
      ['ordered_list', 'bullet_list'],
      [{ heading: ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] }],
      ['code', 'blockquote'],
      ['link', 'image'],
      ['align_left', 'align_center', 'align_right', 'align_justify'],
    ];
    // field must be model bound even if not is use
    this.unused = '';
    
    this.results = [];
    this.pos = -1;
    
    // user(s)
    this.users = [];
    this.display = [];
    
    // schema setup
    this.UserQuery = {
      queryParams: {
        QueryString: '',
        MaximumEntitySuggestions: 10,
        AllowEmailAddresses: true,
        AllowOnlyEmailAddresses: false,
        PrincipalSource: 15,
        PrincipalType: 1,
        SharePointGroupID: 0
      }
    };
    
    this.textKey.pipe(
      debounceTime(250),
      distinctUntilChanged()
    ).subscribe((key) => this.onUpTextSearch(key));
    
    this.userKey.pipe(
      debounceTime(250),
      distinctUntilChanged()
    ).subscribe((key) => this.onUpUserSearch(key));
  }

  // on init, destroy
  ngOnInit(): void {
  }
  ngOnDestroy(): void {
    this.editor.destroy();
  }

  /* 
  Common parts between multiple field types or minimal functions
  */

  // show numbers with only 1 dot and without any trailing zeros
  niceNumber(): string {
    // .toLocaleString() will only retain 3 decimal places therefore split and do dp manually
    // if no dp then no decimal dot either
    // if dp only get 1st, should never be 2 i.e. 0.1.2
    return !this.form[this.field] ? '' : this.form[this.field].toLocaleString().split('.')[0] + (this.form[this.field].toString().split('.').length == 1 ? '' : '.' + this.form[this.field].toString().split('.')[1].replace(/0*$/,''));
  }

  numberSet(e:string|undefined):void {
    if (!e || e == '') {
      this.form[this.field] = null
      return;
    }
    var p = parseFloat(e.replace(/[^0-9\.]/g, ''));
    if (isNaN(p)) {
      this.form[this.field] = null
      return;
    }
    if (p < this.get('Min') && this.get('Min'))
      p = this.get('Min');
    if (p > this.get('Max') && this.get('Max'))
      p = this.get('Max');
    this.form[this.field] = p;
  }

  // field required based on spec but required is not needed for hidden/disabled items
  required(): boolean {
    if (!this.get('Required') || this.disabled || this.elRef.nativeElement.hidden)
      return false;
    return true;
  }

  // max length character countdown
  remaining(): number {
    var m = this.get('MaxLength');
    if (!m && this.get('TypeAsString') == 'Text')
      m = 255;
    if (!m)
      return 15;
    return m - (this.form[this.field] || '').length;
  }

  // gets the required field properties and/or any overrides to determine which field type etc to display
  //declare overrode: any[string];
  get(t:string):any {
    var p:any = null;
    var overrode = this.override ? (typeof this.override == "string" ? JSON.parse(this.override) : this.override) : {};
    if (overrode && overrode[t] != null)
      p = overrode[t];
    if (p == null && this.spec && this.spec[this.field.replace(/^OData_/, '')] && this.spec[this.field.replace(/^OData_/, '')][t])
      p = this.spec[this.field.replace(/^OData_/, '')][t];
    if (p == null && this.spec && this.spec[this.field] && this.spec[this.field][t])
      p = this.spec[this.field][t];
    // if its a multi choice, ensure the object is the correct type
    if (t == 'TypeAsString' && p == 'MultiChoice' && (!this.form[this.field] || !this.form[this.field].results))
      this.form[this.field] = {
        __metadata: {type: "Collection(Edm.String)"},
        results: this.form[this.field] || []
      }
    // if its a multi user, ensure the object is the correct type
    if (t == 'TypeAsString' && p == 'UserMulti' && (!this.form[this.field + 'Id'] || !this.form[this.field + 'Id'].results))
      this.form[this.field + 'Id'] = {
        __metadata: {type: "Collection(Edm.Int32)"},
        results: this.form[this.field + 'Id'] || []
      }
    // if its a url, ensure the correct object type and clone data into url for flat stored occurrences 
    if (t == 'TypeAsString' && p == 'URL' && (!this.form[this.field] || !this.form[this.field].Description))
      this.form[this.field] = {
        Description: '',
        URL: this.form[this.field] || ''
      }
    // if its attachments, ensure the object is the correct type
    if (t == 'TypeAsString' && p == 'Attachments' && (!this.form[this.field] || !this.form[this.field].results))
      this.form[this.field] = {
        results: []
      }
    // if no choices return empty array
    if (t == 'Choices' && p == null)
      return [];
    // if no title use internal field name
    if (t == 'Title' && p == null)
      return this.field;
    if (p == null && (t == 'Min' || t == 'Max') && this.get('TypeAsString') == 'DateTime')
      return (t == 'Min' ? '1970-01-01' : '9999-12-31') + (this.get('DisplayFormat') == 1 ? 'T00:00:00' : '');
    return p == null || typeof p.results == "undefined" ? p : p.results;
  }

  // any field changes trigger for relevant updates
  change(e:any) : void {
    // on multi select, not using ctrl key
    if (this.get('TypeAsString') == 'MultiChoice') {
      e.stopPropagation();
      let scrollTop = 0;
      if ( e.target.parentNode )
        scrollTop = e.target.parentNode.scrollTop;
  
      var v = e.target.innerText;
      // ensure object type is correct
      if (!this.form[this.field] || !this.form[this.field].__metadata)
        this.form[this.field] = {
          __metadata: {type: "Collection(Edm.String)"},
          results: !this.form[this.field] ? [] : this.form[this.field].results ? this.form[this.field].results : typeof this.form[this.field] == "object" ? this.form[this.field] : [this.form[this.field]]
        }
      // if there are selected results set the field to add/remove the most recent click
      var i = this.form[this.field].results.indexOf(v);
      if (i >= 0)
        this.form[this.field].results.splice(i,1);
      else
        this.form[this.field].results.push(v);
  
      const tmp = this.form[this.field].results;
      this.form[this.field].results = [];
      for ( let i = 0; i < tmp.length; i++ ) {
          this.form[this.field].results[i] = tmp[i];
      }
      setTimeout(( function() { e.target.parentNode.scrollTop = scrollTop; } ), 10 );
      setTimeout(( function() { e.target.parentNode.focus(); } ), 10 );
    }
    // append only changes needs 1 way bind to form
    else if (this.get('AppendOnly'))
      this.form[this.field] = e.target ? e.target.value : e;
    // if on change passed in
    if (typeof this.onchange == "function")
      this.onchange(this);
  }
  
  /* 
  Common parts between text fields
  */

  async onUpText(key): Promise<void> {
    if (!this.text.search)
      return;

    if (key == "ArrowDown") {
      if (this.pos < this.results.length - 1)
        this.pos++;
    } else if (key == "ArrowUp") {
      if (this.pos > 0)
        this.pos--;
    } else if (key == "Enter") {
      await this.selectedText(this.results[this.pos]);
    } else {
      if (!this.form[this.field])
        this.results = [];
      else
        this.textKey.next(this.form[this.field]);
    }
  }

  async onUpTextSearch(text:string): Promise<void> {
    if (!this.text.search)
      return;

    this.text.search(this.form[this.field], this.text.parent || this).then((res:any) => {
      this.results = res;
      this.pos = -1;
    });
  }

  async selectedText(res:any): Promise<void> {
    if (!this.text.search)
      return;

    if (!res) {
      this.form[this.field] = null;
      this.results = [];
      return;
    }

    this.form[this.field] = res[this.field];
    if (this.text.select)
      await this.text.select(res, this.text.parent || this);
    this.results = [];
  }

  /* 
  Common parts between choice fields
  */

  // many choices render bigger box
  multiLargeorSmall(): string {
    return this.choices().length > 10 ? 'multilarge' : 'multismall'
  }

  // choices need filtering
  choices(): any[string] {
    // get choices from list
    let choices = this.get('Choices');
    // use any provided filter
    if (typeof this.select.filter == "function")
      choices = choices.filter((c:any, i:number, a:any) => this.select.filter ? this.select.filter(c,i,a,this.select.parent || this) : true);
    // common filters
    var other = this.select.other;
    return choices.filter((x:string) => {
      if (!x || x == '')
        return false;
      // filter exclude other if present
      if (other && other == x)
        return false;
      // exclude unselected items on disabled fields
      if (this.disabled && this.form[this.field] && this.form[this.field].results && !~this.form[this.field].results.indexOf(x))
        return false;
      // filter on search above multichoice field
      if (this.filterMulti && this.filterMulti.length > 0 && !~x.toLowerCase().indexOf(this.filterMulti.toLowerCase()))
        return false;
      // else true
      return true;
    })
  }

  // selected field option not in available choices, i.e. other
  notInChoices(): boolean {
    if (!this.form[this.field] || (this.form[this.field] == '-' && this.select.none))
      return false;
    var choices = this.choices();
    if (!choices)
      return false;
    var ths = this;
    return choices.filter((x:string) => {
        return x == ths.form[ths.field];
      }).length == 0;
  }

  // on single selection change
  selChangeS(v:string): void {
    this.form[this.field] = v;
  }

  // on multi selection change, requires ctrl key
  selChangeM(v:string): void {
    this.form[this.field].results = v;
  }
  
  /* 
  Common parts between file upload field types
  */

  // get outcomes of non standard fields into a plain text field for [required] to be triggered automatically
  attach(): string|undefined {
    return this.attachments().length > 0 ? 'true' : undefined;
  }

  // files post filtering
  attachments(): any[string] {
    if (!this.form[this.field] || !this.form[this.field].results)
      return [];
    var v = this.file.view ?? 0;
    return this.form[this.field].results.filter((f:any) => {
      if (v == 0 || !this.file.archive)
        return true;
      if (v == 1 && !f.ListItemAllFields[this.file.archive])
          return true;
      if (v == -1 && f.ListItemAllFields[this.file.archive])
          return true;
      return false;
    });
  }

  setPrimary(f:any, e:any) {
    // remove primry from all
    this.form[this.field].results.forEach(r => {
      r.Primary = false;
      if (this.file.primary)
        r.ListItemAllFields[this.file.primary] = false;
    });

    // if unchecked then return
    if (!(e.target ? e.target['checked'] : e))
      return;

    // set primary to this
    f.Primary = true;
    if (this.file.primary)
      f.ListItemAllFields[this.file.primary] = true;

    // if it has a doc type, then set primary to all with same doc type
    if (this.file.doctype) {
      this.form[this.field].results.forEach(r => {
        if (!this.file.doctype)
          return;
        r.Primary = r.ListItemAllFields[this.file.doctype] == f.ListItemAllFields[this.file.doctype];
      });
    }
  }

  setClass(f: any, e: any) {
    if (!this.file.doctype)
      return;
    f.ListItemAllFields[this.file.doctype] = e.target ? e.target['value'] : e;
    f.Changed = true;
  }

  width(): string {
    var w = 2;
    this.file.doctypes?.forEach((d) => {
      if (d.length > w)
        w = d.length;
    });
    return `width: ${w}ch`;
  }

  delete(i: number, f?: any, a:boolean = false) {
    if (!f.ServerRelativeUrl) {
      // not uploaded then exclude from potential upload
      this.form[this.field].results.splice(i, 1);
    } else {
      // if uploaded already
      if (!this.file.archive || this.field == 'Attachments') {
        // no archive flag or its attachments so no archiving, then flag for deletion
        if (a && !confirm('Are you sure you wish to delete this file?'))
          return;
        f.Deleted = a;
      } else {
        // toggle archived or delete flag
        if (f.ListItemAllFields[this.file.archive] && (a || f.Deleted)) {
          if (a && !confirm('Are you sure you wish to delete this file?'))
            return;
          f.Deleted = a;
        } else
          f.ListItemAllFields[this.file.archive] = a;
      }
    }

    if (typeof this.onchange == "function")
      this.onchange(this);
  }

  // add attachment to array
  add(file:any) {
    // read the files into the files array
    var files:any = [];
    for (var i = file.files.length - 1; i >= 0; i--)
      files.push(file.files[i]);
    // copy these outside for reuse in the loop
    var ths = this;
    var remaining = files.length;
    // loop the array in forEach for variable isolation
    files.forEach((f:any) => {
      var reader = new window.FileReader();
      reader.onload = async function (event:any) {
        try {
          await ths.appendFile(f.name, event.target.result, ths.form[ths.field].results, false);
          
          if (typeof ths.onchange == "function")
            ths.onchange(this);

          remaining--;
          if (remaining == 0)
            setTimeout(() => file.value = null, 10);
        } catch (e) {
          alert(`File read error: ${f.name} - ${e}`);
        }
      }
      reader.onerror = function (e) {
        alert(`File read error: ${f.name} - ${e}`);
        Logger.log({
          message: `Inside - add(${f.name})`,
          level: LogLevel.Error,
        });
      };
      reader.readAsArrayBuffer(f);
    })
  }

  async outlook (transfer:any) {
    var spc = new SharepointChoiceUtils();

    function mailType(transfer:any, type:string) {
      var item = transfer.getData(type);
      if (item == '')
        return null;
      return JSON.parse(item)
    }

    let maillistrow = mailType(transfer, 'multimaillistmessagerows') || mailType(transfer, 'maillistrow');
    if (maillistrow && maillistrow.mailboxInfos.length > 0) {
      for (var i = 0; i < maillistrow.mailboxInfos.length; i++) {
        var fileName = maillistrow.subjects[i] + ".eml";
        
        try {
          // must double url encode any / in the message id
          var fileContent:string = await spc.callApi(
            App.Tenancy,
            App.GraphClient,
            undefined,
            `https://graph.microsoft.com/v1.0/users/${maillistrow.mailboxInfos[i].mailboxSmtpAddress}/messages/${maillistrow.latestItemIds[i].replace(/\//g, '%252F')}/$value`,
            'GET',
            undefined,
            true
          );
          
          this.appendFile(fileName, new TextEncoder().encode(fileContent).buffer, this.form[this.field].results, false);
        } catch (e) {
          alert(`Email read error: ${fileName} - ${e}`);
        }
      }
    }

    let attachment = mailType(transfer, 'attachment');
    if (attachment) {
      let fileName = attachment.attachmentFile.name;

      // truncate the attachment id 28 chars and = from the end to get the parent message id
      let mail = attachment.attachmentFile.attachmentItemId.substring(0, attachment.attachmentFile.attachmentItemId.length - 29) + "=";

      try {
        // must double url encode any / in the message id and attachment id
        var getAttachment:string = (await spc.callApi(
          App.Tenancy,
          App.GraphClient,
          undefined,
          `https://graph.microsoft.com/v1.0/users/${attachment.mailboxInfo.mailboxSmtpAddress}/messages/${mail.replace(/\//g, '%252F')}/attachments/${attachment.attachmentFile.attachmentItemId.replace(/\//g, '%252F')}`
        )).contentBytes;

        this.appendFile(fileName, Uint8Array.from(atob(getAttachment), c => c.charCodeAt(0)).buffer, this.form[this.field].results, false);
      } catch (e) {
        alert(`Attachment read error: ${fileName} - ${e}`);
      }
    }
  }

  async appendFile(fileName:string, data:ArrayBuffer, results:any, skip:boolean) {
    // skip small images
    if (~fileName.toString().toLowerCase().indexOf(".png")
      || ~fileName.toString().toLowerCase().indexOf(".jpg")
      || ~fileName.toString().toLowerCase().indexOf(".gif"))
      if (skip && data.byteLength < 8096)
        return;

    // cleanup the name, more agressive as it may not be from a windows file system
    var n = fileName.trim().replace(/[\\/:*?"%'#<>|]/g,'-');
    // get the extension
    var e = n.substring(n.lastIndexOf('.')+1);
    // get the first part of the name
    var f = n.substring(0, n.lastIndexOf('.'));
    // get the title
    var t = f.length > 255 ? f.substring(0, 255) : f;
    // get shortened file name
    var s = f.length > 100 ? f.substring(0, 100) : f;

    // find the next available name by appending a number
    var i = 1, newName = `${s}.${e}`;
    while (results.filter(f => f.FileName == newName).length > 0) {
      newName = `${s} (${i++}).${e}`;
    }
    
    results.push({
      FileName: newName,
      Data: data,
      Length: data.byteLength,
      ListItemAllFields: { Title: t }
    });
    
    if (!this.file.extract)
      return;

    if (fileName.toLowerCase().endsWith(".zip"))
      await this.zips(data, results);
    if (fileName.toLowerCase().endsWith(".msg"))
      await this.msgs(data, results);
    if (fileName.toLowerCase().endsWith(".eml"))
      await this.emls(data, results);
  }

  // extract zip files and append to results
  async zips(data:ArrayBuffer, results:Array<any>) {
    try {
      var zip = await JSZip.loadAsync(data);
      var files = Object.keys(zip.files);
      files.forEach(async (file) => {
        try {
          var buffer:ArrayBuffer|undefined = await zip.file(file)?.async('arraybuffer');
          if (buffer)
            await this.appendFile(file, buffer, results, false);
        } catch (e) { }
      });
    } catch (e) {
      // zip is uploaded so any extracted elements are only nice to have
    }
  }

  // extract and append msg email attachments to results
  async msgs(data:ArrayBuffer, results:Array<any>) {
    try {
      var msgReader = new MsgReader.MSGReader(data);
      // needs to be triggered to get the parser
      msgReader.getFileData();
      var i = 0;
      // keep going until error because the part of this module that gives the count isnt mapped in typescript
      while (true) {
        var file = msgReader.getAttachment(i++);
        try {
          // square brackets for buffer here as file.content is a Uint8Array but the typing shows it as string
          await this.appendFile(file.fileName, file.content['buffer'], results, true);
        } catch (e) { }
      }
    } catch (e) {
      // msg is uploaded so any extracted elements are only nice to have
    }
  }

  // extract and append eml email attachments to results
  async emls(data:ArrayBuffer, results:Array<any>) {
    try {
      // reads the email string data into a json object
      readEml(new TextDecoder().decode(data), (err, ReadEmlJson) => {
        if (err || !ReadEmlJson || !ReadEmlJson.attachments)
          return;
        ReadEmlJson.attachments.forEach(async (attachment:any) => {
          try {
            await this.appendFile(attachment.name, Uint8Array.from(atob(attachment.data64), c => c.charCodeAt(0)).buffer, results, true);
          } catch (e) { }
        });
      });
    } catch (e) {
      // eml is uploaded so any extracted elements are only nice to have
    }
  }

  // dragging and dropping, hover
  over(evt:any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.filesOver = true;
  }

  // dragging and dropping, unhover
  leave(evt:any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.filesOver = false;
  }
  
  // dragging and dropping, drop
  drop(evt:any) {
    evt.preventDefault();
    evt.stopPropagation();
    if (evt.dataTransfer?.files?.length > 0)
      this.add(evt.dataTransfer);
    if (evt.dataTransfer?.items?.length > 0)
      this.outlook(evt.dataTransfer);
    this.filesOver = false;

    // if no transfer on drop and not chromium based add the meta tag to allow the drop next time
    if (!evt.dataTransfer && !window['chrome']) {
      var m = document.createElement("meta");
      m.httpEquiv = "X-UA-Compatible";
      m.content = "chrome=1";
      document.head.appendChild(m);
    }
  }

  /* 
  Common parts between user field types
  */

  // get outcomes of non standard fields into a plain text field for [required] to be triggered automatically
  people(): string|undefined {
    return this.form[this.field + 'Id'] && (!this.form[this.field + 'Id'].results || this.form[this.field + 'Id'].results.length > 0) ? 'true' : undefined;
  }
  
  // select user
  async selectedUser(res:any): Promise<void> {
    if (!this.spec['odata.context'])
      return;

    // ensure correct schema
    if (this.get('TypeAsString') == 'UserMulti' && (!this.form[this.field + 'Id'] || !this.form[this.field + 'Id'].__metadata))
      this.form[this.field + 'Id'] = {
        __metadata: {type: "Collection(Edm.Int32)"},
        results: !this.form[this.field + 'Id'] ? [] : this.form[this.field + 'Id'].results ? this.form[this.field + 'Id'].results : typeof this.form[this.field + 'Id'] == "object" ? this.form[this.field + 'Id'] : [this.form[this.field + 'Id']]
      }
    if (this.get('TypeAsString') == 'User' && this.form[this.field + 'Id'] && this.form[this.field + 'Id'].results)
      this.form[this.field + 'Id'] = this.form[this.field + 'Id'].results.length > 0 ? this.form[this.field + 'Id'].results[0] : null;

    // use click item
    if (res) {
      // already selected, do nothing
      if (this.get('TypeAsString') == 'UserMulti' && this.display.filter((x:any) => {
        return x.Key == res.Key
      }).length > 0)
        return;
      if (this.get('TypeAsString') == 'User' && this.display.length > 0 && this.display[0].Key == res.Key)
        return;

      // setup context late to adapt to changes
      var u = await this.spec['odata.context'].web.ensureUser(res.Key);

      // add to field
      if (this.get('TypeAsString') == 'UserMulti') {
        this.form[this.field + 'Id'].results.push(u.Id);
        this.display.push({
          DisplayText: res.DisplayText,
          Key: res.Key,
          Id: u.Id
        });
      } else {
        this.form[this.field + 'Id'] = u.Id;
        this.display = [{
          DisplayText: res.DisplayText,
          Key: res.Key,
          Id: u.Id
        }];
      }
    }

    // clear search fields
    this.name = '';
    this.users = [];
    
    if (typeof this.onchange == "function")
      this.onchange(this);
  }

  // load list data only has IDs so expand the object
  displayUser(user:any): string {
    if (!this.spec['odata.context'] || !user)
      return user || '';

    var u = this.display.filter((x:any) => {
      return x.Id == user
    });
    if (u.length > 0 && u[0].DisplayText)
      return u[0].DisplayText;

    // setup context late to adapt to changes
    if (!this.loading)
      this.loading = [];
    
    // dont trigger a new load web request if the users aready loading
    if (this.loading.indexOf(user) < 0 && typeof user == "number" && user > 0) {
      this.loading.push(user);
      // load the user
      this.spec['odata.context'].web.getUserById(user)().then((u:any) => {
        // update the display table
        this.display.push({
          DisplayText: u.Title,
          Key: u.LoginName,
          Id: u.Id
        });
        // touch results to force display update
        if (this.form[this.field + 'Id'].results) {
          this.form[this.field + 'Id'].results.push(0);
          this.form[this.field + 'Id'].results.pop();
        }
      });
    }
    
    return '';
  }

  // removes a user
  removeUser(usr:any): void {
    if (!usr) {
      this.form[this.field + 'Id'] = null;
      this.display = [];
      return;
    }
    
    this.form[this.field + 'Id'].results.splice(this.form[this.field + 'Id'].results.indexOf(usr), 1);
    this.display = this.display.filter((x:any) => {
      return x.Id != usr
    })

    if (typeof this.onchange == "function")
      this.onchange(this);
  }

  // trigger user search
  async onUpUser(key): Promise<void> {
    if (key == "ArrowDown") {
      if (this.pos < this.users.length - 1)
        this.pos++;
    } else if (key == "ArrowUp") {
      if (this.pos > 0)
        this.pos--;
    } else if (key == "Enter") {
      this.selectedUser(this.users[this.pos]);
    } else {
      if (!this.name || this.name.length < 3)
        this.users = [];
      else
        this.userKey.next(this.name);
    }
  }

  async onUpUserSearch(text:string): Promise<void> {
    if (!this.spec['odata.context'])
      return;

    var url = (await this.spec['odata.context'].web()).ServerRelativeUrl;

    // set the user partial being searched
    this.UserQuery.queryParams.QueryString = this.name;
    // set group each query to adapt to changes
    this.UserQuery.queryParams.SharePointGroupID = parseInt(this.get('SelectionGroup') || 0);
    // ensure up to date digest for http posting
    var token:any = await fetch(url + '/_api/contextinfo', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json;odata=verbose'
      }
    });
    let digest = await token.json();
    // query users api, no pnp endpoint for this
    var search:any = await fetch(url + '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': digest.d.GetContextWebInformation.FormDigestValue
      },
      body: JSON.stringify(this.UserQuery)
    });
    let res:any = await search.json();
    this.users = [];
    const allUsers: User[] = JSON.parse(res.d.ClientPeoplePickerSearchUser);
    allUsers.filter(x => {
      return x.EntityData.Email && !~x.Key.indexOf('_adm') && !~x.Key.indexOf('adm_')
    }).forEach(user => {
      this.users = [...this.users, user];
    });
  }
}
