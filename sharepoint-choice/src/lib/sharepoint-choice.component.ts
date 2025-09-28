import { Component, OnInit, OnDestroy, Input, ElementRef, ChangeDetectorRef, ErrorHandler } from '@angular/core';
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import { Editor, NgxEditorModule, Toolbar } from 'ngx-editor';
import MsgReader from '@kenjiuno/msgreader';
import { readEml } from 'eml-parse-js';
import { loadAsync } from 'jszip';
import { Subject } from 'rxjs';
import { debounceTime, distinctUntilChanged } from 'rxjs/operators';
import { SharepointChoiceUtils } from './sharepoint-choice.utils';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SharepointChoiceLogging } from './sharepoint-choice.logging';

@Component({
  selector: 'app-choice',
  templateUrl: './sharepoint-choice.component.html',
  styleUrls: ['../styles.scss'],
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    NgxEditorModule
  ],
  providers: [{
    provide: ErrorHandler,
    useClass: SharepointChoiceLogging
  }]
})
export class SharepointChoiceComponent implements OnInit, OnDestroy {
  @Input() form!: object; // form containing field, varies based on list therefore object not defined explicitly
  @Input() field!: string; // internal field name on form object, used for push back and against spec

  @Input() spec!: object; // spec of field loaded from list, varies based on list therefore object not defined explicitly
  @Input() versions!: Array<object>; // version history of this field to display if presented, varies based on list therefore object not defined explicitly
  @Input() override!: string|object; // manually override any spec above. prefer send as string as passing object kills large form performance

  @Input() disabled!: boolean; // get disabled state from outside
  @Input() onchange!: Function; // onchange trigger a function(this)

  @Input() text!: { // override text for field
    pattern?: string, // regex pattern for validation
    height?: number, // height of text area in px

    search?: Function, // search via api for drop down options
    select?: Function, // upon selection in drop down call back function
    parent?: any // parent object that the control belongs to for call backs
  };

  @Input() select!: { // override select for field
    none?: string, // none option text instead of null
    other?: string, // Other fill-in option text, will override to allow other

    filter?: Function // filter choices by a function
    parent?: any // parent object that the control belongs to for call backs
  };

  @Input() file!: { // override file for field
    extract?: boolean, // extract files from zip and email
    check?: boolean, // show check box for each file

    accept?: string, // accept file types attribute
    download?: boolean, // force download of files
    uploadonly?: boolean, // only upload files

    archive?: string, // archive field name
    view?: number, // view type 0 - all, 1 - not archived, -1 - archived

    doctypes?: Array<string>, // document types
    doctype?: string, // document type field name

    notes?: string, // notes input field name for singular note input space
    spec?: object // field spec for additional fields, 
  };

  declare editor?: Editor;
  declare toolbar: Toolbar;
  declare tooltip: boolean;
  declare filesOver: boolean;
  declare name: string;
  declare loading: Array<number>;
  declare versionsDisplayed: boolean;

  declare users: Array<{
    Key: string;
    Description: string;
    DisplayText: string;
    EntityType: string;
    ProviderDisplayName: string;
    ProviderName: string;
    IsResolved: boolean;
    EntityData: {
      IsAltSecIdPresent: string;
      Title: string;
      Email: string;
      MobilePhone: string;
      ObjectId: string;
      Department: string;
    };
    MultipleMatches: any[];
  }>;

  declare filterMulti: string;
  declare unused: string;
  declare results: Array<object>; // varies based on the search source therefore not explicitly defined
  declare pos: number;
  declare office: {
    type: string | null,
    loading: boolean
  };

  declare sort: string;
  declare filter: string;

  public textKey: Subject<string> = new Subject<string>();
  public userKey: Subject<string> = new Subject<string>();

  private overridePrevious?: string;
  private overrideParsed?: object;

  private display: Array<{
    Id: number;
    Key: string;
    DisplayText: string;
  }>;
  
  // drop external models
  private UserQuery: {
    queryParams: {
      QueryString: string;
      MaximumEntitySuggestions: number;
      AllowEmailAddresses: boolean;
      AllowOnlyEmailAddresses: boolean;
      PrincipalType: number;
      PrincipalSource: number;
      SharePointGroupID: number;
    };
  };

  constructor(
    private elRef: ElementRef,
    private chRef: ChangeDetectorRef
  ) {
    this.office = {
      type: null,
      loading: false
    };
    if (!this.text)
      this.text = {};
    if (!this.select)
      this.select = {};
    if (!this.file)
      this.file = {};
    if (!this.filter)
      this.filter = '';

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
    this.editor?.destroy();
  }

  /* 
  Common parts between multiple field types or minimal functions
  */

  // show or hide tooltips 
  showHideTooltip(show: boolean): void {
    this.tooltip = show;
    this.chRef.detectChanges();
  }

  // are there different field version values shown
  versionsToggle(): string {
    if (!this.versionsDisplayed) {
      this.versionsDisplayed = true;
      this.chRef.detectChanges();
    }
    return '';
  }

  // show numbers with only 1 dot and without any trailing zeros
  niceNumber(): string {
    // .toLocaleString() will only retain 3 decimal places therefore split and do dp manually
    // if no dp then no decimal dot either
    // if dp only get 1st, should never be 2 i.e. 0.1.2
    if (!this.form[this.field] && this.form[this.field] !== 0)
      return '';
    var s = this.form[this.field].toLocaleString().split('.');
    return s[0] + (s.length == 1 ? '' : '.' + s[1].replace(/0*$/, ''));
  }

  numberSet(e: string | undefined): void {
    if (!e) {
      this.form[this.field] = null
      return;
    }
    var p = parseFloat(e.replace(/[^0-9\.]/g, ''));
    if (isNaN(p)) {
      this.form[this.field] = null
      return;
    }
    var min = this.get('Min');
    if (min != null && p < min)
      p = min;
    var max = this.get('Max');
    if (max != null && p > max)
      p = max;
    this.form[this.field] = p;
  }

  // field required based on spec but required is not needed for hidden/disabled items
  required(): boolean {
    if (this.disabled || this.elRef.nativeElement.hidden || !this.get('Required'))
      return false;
    return true;
  }

  // get outcomes of non standard fields into a plain text field for [required] to be triggered automatically 
  validator(): string {
    switch (this.get('TypeAsString')) {
      case 'User':
        return this.form[this.field + 'Id'] ? 'true' : '';
      case 'UserMulti':
        return (this.form[this.field + 'Id'] && this.form[this.field + 'Id'].results &&  this.form[this.field + 'Id'].results.length > 0) ? 'true' : '';
      case 'Attachments':
        return this.attachments().length > 0 ? 'true' : '';
      case 'Choice':
        return this.form[this.field] ? 'true' : '';
      default:
        // these will fall back on default html required validation
        return 'true';
    }
  }

  // max length character countdown
  remaining(): number {
    var m = this.get('MaxLength');
    if (!m && this.get('TypeAsString') == 'Text')
      m = 255;
    if (!m)
      return 255;
    return m - (this.form[this.field] || '').length;
  }

  // gets the required field properties and/or any overrides to determine which field type etc to display
  //declare overrode: any[string];
  get(t: string): any {
    // if override passed in, make sure its an object and convert only once unless changed
    if (this.override && typeof this.override === 'string' && this.override != this.overridePrevious) {
      this.overridePrevious = this.override;
      this.overrideParsed = JSON.parse(this.override);
    }
    // if override is an object then use it directly
    if (this.override && typeof this.override !== 'string') {
      this.overrideParsed = this.override;
    }
    
    // initial starting value of p from override
    var p: any = this.overrideParsed ? this.overrideParsed[t] : null;
    // if no override then get from field spec selected
    if (p == null) {
      let spec = this.spec[this.field.replace(/^OData_/, '')] ?? this.spec[this.field];
      if (spec)
        p = spec[t];
    }

    if (t == 'TypeAsString') {
      // if its a multi choice, ensure the object is the correct type
      if (p == 'MultiChoice') {
        if (!this.form[this.field] || !this.form[this.field].results)
          this.form[this.field] = {
            __metadata: { type: "Collection(Edm.String)" },
            results: this.form[this.field] || []
          }
      // if its a multi user, ensure the object is the correct type
      } else if (p == 'UserMulti') {
        if (!this.form[this.field + 'Id'] || !this.form[this.field + 'Id'].results)
          this.form[this.field + 'Id'] = {
            __metadata: { type: "Collection(Edm.Int32)" },
            results: this.form[this.field + 'Id'] || []
          }
      // if its a url, ensure the correct object type and clone data into url for flat stored occurrences 
      } else if (p == 'URL') {
        if (!this.form[this.field] || !this.form[this.field].Description)
          this.form[this.field] = {
            Description: '',
            URL: this.form[this.field] || ''
          }
      // if its attachments, ensure the object is the correct type and office fired if needed
      } else if (p == 'Attachments') {
        this.officeAddin();
        if (!this.form[this.field] || !this.form[this.field].results)
          this.form[this.field] = {
            results: []
          }
      }
      // if the field is empty (not set null) and its not a person field (+Id) then set the default value only if there is one
      if (this.form[this.field] === undefined && !this.form[this.field + 'Id']) {
        var d = this.get('DefaultValue');
        if (d)
          this.form[this.field] = d;
      }
      // always disable read only fields initially, may get overridden later but thats the consumers responsibility
      var r = this.get('ReadOnlyField');
      if (r)
        this.disabled = true;
    }

    // init rich text editor
    else if (t == 'RichText' && p && !this.editor)
      this.editor = new Editor();

    // if no choices return empty array
    else if (t == 'Choices' && p == null)
      return [];
    // if no title use internal field name
    else if (t == 'Title' && p == null)
      return this.friendlyName(this.field);
    // min max of date time fields
    else if ((t == 'Min' || t == 'Max') && p == null && this.get('TypeAsString') == 'DateTime')
      return (t == 'Min' ? '1970-01-01' : '9999-12-31') + (this.get('DisplayFormat') == 1 ? 'T00:00:00' : '');
    
    return p == null || !p.results ? p : p.results;
  }

  // any field changes trigger for relevant updates
  async change(e: any): Promise<void> {
    // on multi select, not using ctrl key
    if (this.get('TypeAsString') == 'MultiChoice') {
      e.stopPropagation();
      let scrollTop = 0;
      if (e.target.parentNode)
        scrollTop = e.target.parentNode.scrollTop;

      var v = e.target.innerText;
      // ensure object type is correct
      if (!this.form[this.field] || !this.form[this.field].__metadata)
        this.form[this.field] = {
          __metadata: { type: "Collection(Edm.String)" },
          results: !this.form[this.field] ? [] : this.form[this.field].results ? this.form[this.field].results : typeof this.form[this.field] == "object" ? this.form[this.field] : [this.form[this.field]]
        }
      // if there are selected results set the field to add/remove the most recent click
      var i = this.form[this.field].results.indexOf(v);
      if (i >= 0)
        this.form[this.field].results.splice(i, 1);
      else
        this.form[this.field].results.push(v);

      const tmp = this.form[this.field].results;
      this.form[this.field].results = [];
      for (let i = 0; i < tmp.length; i++) {
        this.form[this.field].results[i] = tmp[i];
      }
      setTimeout((function () { e.target.parentNode.scrollTop = scrollTop; }), 10);
      setTimeout((function () { e.target.parentNode.focus(); }), 10);
    }
    // append only changes needs 1 way bind to form
    else if (this.get('AppendOnly'))
      this.form[this.field] = e.target ? e.target.value : e;
    // if on change passed in
    if (typeof this.onchange == "function")
      await this.onchange(this);

    this.chRef.detectChanges();
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

  async onUpTextSearch(text: string): Promise<void> {
    if (!this.text.search)
      return;

    this.results = await this.text.search(this.form[this.field], this.text.parent || this);

    this.pos = -1;
    this.chRef.detectChanges();
  }

  async selectedText(res: any): Promise<void> {
    if (!this.text.search)
      return;

    if (!res) {
      this.form[this.field] = null;
    } else {
      this.form[this.field] = res[this.field];
      if (this.text.select)
        await this.text.select(res, this.text.parent || this);
    }

    this.results = [];
    this.chRef.detectChanges();
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
      choices = choices.filter((c: any, i: number, a: any) => this.select.filter ? this.select.filter(c, i, a, this.select.parent || this) : true);
    // common filters
    return choices.filter((x: string) => {
      // exclude unselected items on disabled fields
      if (this.disabled && this.form[this.field] && this.form[this.field].results && !~this.form[this.field].results.indexOf(x))
        return false;
      // filter on search above multichoice field
      if (this.filterMulti && !~x.toLowerCase().indexOf(this.filterMulti.toLowerCase()))
        return false;
      // else true
      return true;
    })
  }

  // selected field option not in available choices, i.e. other
  inChoices(): boolean {
    if (!this.form[this.field])
      return true;
    return this.choices().indexOf(this.form[this.field]) >= 0;
  }

  // on single selection change
  selChangeS(v: string): void {
    this.form[this.field] = v;
  }

  // on multi selection change, requires ctrl key
  selChangeM(v: string): void {
    this.form[this.field].results = v;
  }

  /* 
  Common parts between file upload field types
  */

  // files post filtering
  attachments(): any[string] {
    if (!this.form[this.field] || !this.form[this.field].results)
      return [];
    var v = this.file.view ?? 0;
    return this.form[this.field].results
      .filter((f: any) => {
        if (v == 0 || !this.file.archive)
          return true;
        if (v == 1 && !f.ListItemAllFields[this.file.archive])
          return true;
        if (v == -1 && f.ListItemAllFields[this.file.archive])
          return true;
        return false;
      })
      .filter((f: any) => {
        if (!this.filter || this.filter == '' || !this.file.doctype)
          return true;
        if (!f.ListItemAllFields || !f.ListItemAllFields[this.file.doctype])
          return true;
        if (f.ListItemAllFields[this.file.doctype] == this.filter)
          return true;
        return false;
      })
      .sort((a: any, b: any) => {
        if (!this.sort || this.sort == '' || this.sort == '-')
          this.sort = 'Created';
        var s = this.sort.replace(/^[-\+]/, '');
        var o = this.sort.startsWith('-') ? -1 : 1;
        if (!a.ListItemAllFields || !a.ListItemAllFields[s])
          return -2;
        if (a.ListItemAllFields[s] < b.ListItemAllFields[s])
          return o;
        if (a.ListItemAllFields[s] > b.ListItemAllFields[s])
          return -o;
        return 0;
      });
  }

  friendlyName(name: string): string {
    return name.replace(/_x0020_/g, ' ').replace(/([a-z])([A-Z])/g, '$1 $2');
  }

  hasChecked(): boolean {
    if (!this.form[this.field] || !this.form[this.field].results)
      return false;
    // check if there are any files that have been checked
    return this.form[this.field].results.some((f: any) => f.Checked);
  }

  setClasses(e: any): void {
    // set the class of the file
    this.form[this.field].results.forEach(f => {
      if (f.Checked) {
        f.ListItemAllFields['DocumentType'] = e.target.value;
        f.Checked = false;
      }
    });
    e.target.value = null;
  }

  changeSort(field: string) {
    if (this.sort == '-' + field) {
      this.sort = '+' + field;
    } else if (this.sort == '+' + field) {
      this.sort = '';
    } else {
      this.sort = '-' + field;
    }
  }

  additionalKeys(o): Array<string> {
    // get keys of object
    if (!o || typeof o != "object")
      return [];
    return Object.keys(o).filter(k => k != this.file?.doctype && k != this.file?.notes);
  }

  loadOrGenerateSpec(field:string, type:string) {
    var s:object = {};
    if (this.file?.spec && this.file.spec[field])
       s[field] = this.file.spec[field];
    else
      s[field] = { 
        TypeAsString: type,
        InternalName: field,
        Title: '',
        Choices: this.file?.doctypes ?? []
      };
    return s;
  }

  usedTypes(): Array<string> {
    // get initial types
    var types = this.file.doctypes || [];
    // get all types used in the attachments
    if (this.file.doctype && this.form[this.field] && this.form[this.field].results)
      types = this.form[this.field].results
        .map((f: any) => f.ListItemAllFields && this.file.doctype && this.file.doctype in f.ListItemAllFields ? f.ListItemAllFields[this.file.doctype] : null)
        .filter((f: any) => f)
        .sort();
    // remove duplicates
    types = [...new Set(types)];
    return types;
  }

  newTab(f, e) {
    window.open(`${f.ServerRelativeUrl || '#'}?Web=1`, '_blank');
  }

  width(): string {
    var c = 0;
    for (var i = 0; i < this.file.doctypes.length; i++) {
      let l = this.file.doctypes[i].length
      if (l > c)
        c = l;
    }
    if (c == 0)
      return '';
    return `width: ${c}ch`;
  }

  async delete(f?: any, a: boolean = false) {
    if (!f.ServerRelativeUrl) {
      // not uploaded then exclude from potential upload
      this.form[this.field].results = this.form[this.field].results.filter((x: any) => x.FileName != f.FileName);
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
      await this.onchange(this);

    this.chRef.detectChanges();
  }

  // add attachment to array
  async add(file: any) {
    // read the files into the files array
    var files: any = [];
    for (var i = file.files.length - 1; i >= 0; i--)
      files.push(file.files[i]);
    // copy these outside for reuse in the loop
    var ths = this;
    var remaining = files.length;
    // loop the array in forEach for variable isolation
    files.forEach((f: any) => {
      var reader = new window.FileReader();
      reader.onload = async function (event: any) {
        try {
          await ths.appendFile(f.name, event.target.result, ths.form[ths.field].results);

          if (typeof ths.onchange == "function") {
            await ths.onchange(this);
            ths.chRef.detectChanges();
          }

          remaining--;

          // if last file added then clear the file input
          if (remaining == 0)
            setTimeout(() => {
              file.value = null;
              ths.chRef.detectChanges();
            }, 10);
        } catch (e) {
          alert(`File onread error: ${f.name} with error ${e}`);
        }
      }
      reader.onerror = function (e) {
        alert(`File read onerror: ${f.name} with error ${e}`);
        throw e;
      };
      // may need to consider how to await these each until all done if extractions start getting timely
      reader.readAsArrayBuffer(f);
    })
  }

  // gets a drag and drop new outlook item which includes ids not file data and adds to the files array
  async outlook(transfer: any) {
    var spc = new SharepointChoiceUtils();

    function mailType(transfer: any, type: string) {
      var item = transfer.getData(type);
      if (!item)
        return null;
      return JSON.parse(item)
    }

    // one or more email messages dropped
    let maillistrow = mailType(transfer, 'multimaillistmessagerows') || mailType(transfer, 'maillistrow');
    if (maillistrow) {
      var errors: Array<string> = [];
      for (var i = 0; i < maillistrow.mailboxInfos.length; i++) {
        var fileName = maillistrow.subjects[i].trim() + ".eml";

        try {
          // must double url encode any / in the message id
          var fileContent: string = await spc.callApi(
            undefined,
            undefined,
            undefined,
            `https://graph.microsoft.com/v1.0/users/${maillistrow.mailboxInfos[i].mailboxSmtpAddress}/messages/${maillistrow.latestItemIds[i].replace(/\//g, '%252F')}/$value`,
            'GET',
            undefined,
            'text'
          );

          await this.appendFile(fileName, new TextEncoder().encode(fileContent).buffer as ArrayBuffer, this.form[this.field].results);
        } catch (e) {
          errors.push(`Email append error: ${fileName} with error ${e}`);
        }
      }

      if (errors.length > 0) {
        // if there are errors then alert them
        alert(`Errors saving emails:\n\n${errors.join('\n')}`);
        throw errors.join('\n');
      }
    }

    // an attachment dropped
    let attachment = mailType(transfer, 'attachment');
    if (attachment) {
      let fileName = attachment.attachmentFile.name;

      // truncate the attachment id 28 chars and = from the end to get the parent message id
      let mail = attachment.attachmentFile.attachmentItemId.substring(0, attachment.attachmentFile.attachmentItemId.length - 29) + "=";

      try {
        // must double url encode any / in the message id and attachment id
        var getAttachment: any = await spc.callApi(
          undefined,
          undefined,
          undefined,
          `https://graph.microsoft.com/v1.0/users/${attachment.mailboxInfo.mailboxSmtpAddress}/messages/${mail.replace(/\//g, '%252F')}/attachments/${attachment.attachmentFile.attachmentItemId.replace(/\//g, '%252F')}`
        );

        await this.appendFile(fileName, Uint8Array.from(atob(getAttachment.contentBytes), c => c.charCodeAt(0)).buffer, this.form[this.field].results, `Sent: ${new Date(getAttachment.lastModifiedDateTime)}`);
      } catch (e) {
        alert(`Error saving attachment: ${fileName} with error ${e}`);
        throw e;
      }
    }

    // a teams file drop, only works for teams libraries not onedrive/chat
    let spo = mailType(transfer, 'application/x-item-keys');
    if (spo) {
      var errors: Array<string> = [];
      for (var i = 0; i < spo.itemKeys.length; i++) {
        try {
          // the inner is still JSON encoded from teams
          spo.itemKeys[i] = JSON.parse(spo.itemKeys[i]);

          var web = Web([spc.sp.web, spo.itemKeys[i][1]]);
          var folder = await web.getFolderByServerRelativePath(spo.itemKeys[i][2].substring(spo.itemKeys[i][2].indexOf('/', 9))).properties();

          var list = folder['vti_x005f_listtitle'] || folder['vti_listtitle'] || folder['listtitle'] || folder['title'];
          var item = await web.lists.getByTitle(list).items.getById(spo.itemKeys[i][3]).select('File').expand('File')();
          var desc = `Created: ${new Date(item.File.TimeCreated)} - Modified: ${new Date(item.File.TimeLastModified)}`;

          var buffer = await web.getFileByServerRelativePath(item.File.ServerRelativeUrl).getBuffer();
          await this.appendFile(item.File.Name, buffer, this.form[this.field].results, desc);
        } catch (e) {
          errors.push(`SharePoint file append error: ${spo.itemKeys[i][2]} with error ${e}`);
        }
      }

      if (errors.length > 0) {
        // if there are errors then alert them
        alert(`Errors saving SharePoint files:\n\n${errors.join('\n')}`);
        throw errors.join('\n');
      }
    }
  }

  officeAddin(): void {
    // dont double load the script
    if (document.getElementById('officejs') || window['Office'])
      return;
    // try and determine if we are in an office addin
    try {
      if (window.top != window.self) {
        throw "In an iframe";
      }
      if ('IsOfficeURLSchemes' in window) {
        throw "In an addin";
      }
      // unlikely to be an addin
      return;
    } catch (e) {
      // continue as its probably an addin
    }
    // load the office.js script as web/pwa apps are iframes and full clients are not with no obvious way to detect if its needed
    var s = document.createElement('script');
    s.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
    s.id = 'officejs';
    document.head.appendChild(s);
    // capture the office type for later use along with triggering change detection if needed
    var ths = this;
    s.addEventListener('load', () => {
      // only once office.js loaded
      if ('Office' in window) {
        var Office: any = window['Office'];
        // on ready should be ran soon after office.js is loaded trigger
        Office.onReady(info => {
          // capture what office type of addin for later use
          if (ths.office.type != info.host) {
            ths.office.type = info.host;
            ths.chRef.detectChanges();
          }
        });
      }
    }, false);
  }

  // import from office addin selection or document panel
  importOutlook() {
    this.office.loading = true;
    var Office: any = window['Office'];

    // if the adding type is outlook then get the selected email(s)
    var spc = new SharepointChoiceUtils();
    var ths = this;

    if (Office.context.mailbox.initialData.isFromSharedFolder) {
      Office.context.mailbox.item.getSharedPropertiesAsync(shared => {
        if (shared.status === Office.AsyncResultStatus.Failed)
          return;
        this.getMailItem(Office, shared?.value?.targetMailbox, spc, ths);
      });
    } else {
      this.getMailItem(Office, null, spc, ths);
    }
  }

  getMailItem(Office:any, targetMailbox:string|null, spc:SharepointChoiceUtils, ths:SharepointChoiceComponent) {
    Office.context.mailbox.getSelectedItemsAsync(async (asyncResult: any) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed)
        return;

      var errors: Array<string> = [];
      for (var m in asyncResult.value) {
        try {
          var message = asyncResult.value[m];

          // must double url encode any / in the message id
          var fileContent: string = await spc.callApi(
            undefined,
            undefined,
            undefined,
            targetMailbox == null ?
              `https://graph.microsoft.com/v1.0/me/messages/${message.itemId.replace(/\//g, '%252F')}/$value` :
              `https://graph.microsoft.com/v1.0/users/${targetMailbox}/messages/${message.itemId.replace(/\//g, '%252F')}/$value`,
            'GET',
            undefined,
            'text'
          );

          await ths.appendFile(`${message.subject.trim()}.eml`, new TextEncoder().encode(fileContent).buffer as ArrayBuffer, ths.form[ths.field].results);
        } catch (e) {
          errors.push(`Email append error: ${message.subject} with error ${e}`);
        }
      }

      if (errors.length > 0) {
        // if there are errors then alert them
        alert(`Errors saving emails:\n\n${errors.join('\n')}`);
        throw errors.join('\n');
      }
    });
  }

  // import from office addin selection or document panel
  importOffice() {
    this.office.loading = true;
    var Office: any = window['Office'];

    // if the adding type is word or excel then get the current document
    var docDataSlices: any = [];
    var slicesReceived = 0;
    var file: any;
    var ths = this;

    try {
      // get the file in 64k slices until complete
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (asyncResult: any) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed)
          return;

        file = asyncResult.value;
        try {
          file.getSliceAsync(0, processSlice);
        } finally {
          file.closeAsync();
        }
      });

      // process the slice and append until completed
      function processSlice(sliceResult: any) {
        if (sliceResult.status === Office.AsyncResultStatus.Failed)
          throw sliceResult.error.message;

        docDataSlices[sliceResult.value.index] = sliceResult.value.data;
        if (++slicesReceived == file.sliceCount)
          saveFile();
        else
          file.getSliceAsync(slicesReceived, processSlice);
      }

      async function saveFile() {
        // completed getting slices then start patching the file
        let docData = [];
        for (let i = 0; i < docDataSlices.length; i++)
          docData = docData.concat(docDataSlices[i]);

        // convert char codes to string
        let fileContent = new String();
        for (let j = 0; j < docData.length; j++)
          fileContent += String.fromCharCode(docData[j]);

        // try to get a file name from the document
        var fileName = Office.context.document.url?.split('/').pop().split('\\').pop();
        if (!fileName || fileName == '') {
          fileName = `OfficeAddin-${new Date().toISOString().replace(/:/g, '-')}`;
          switch (ths.office.type) {
            case "Word":
              fileName += ".docx";
              break;
            case "Excel":
              fileName += ".xlsx";
              break;
            case "PowerPoint":
              fileName += ".pptx";
              break;
            case "OneNote":
              fileName += ".one";
              break;
            case "Project":
              fileName += ".mpp";
              break;
            case "Visio":
              fileName += ".vsdx";
              break;
            default:
              throw "Unknown Office type, unable to save file";
          }
        }

        await ths.appendFile(`${fileName}`, Uint8Array.from(fileContent.toString(), c => c.charCodeAt(0)).buffer, ths.form[ths.field].results);
      }
    } catch (e) {
      // if there is an error then alert it
      alert(`Error saving file:\n\n${e}`);
      throw e;
    }
  }

  async appendFile(fileName: string, data: ArrayBuffer, results: any, desc?: string) {
    // only save files that have an extension
    if (fileName.indexOf('.') < 0)
      return;
    // cleanup the name, more agressive as it may not be from a windows file system or have trailing chars from email systems
    var n = fileName.trim().replace(/[\\/:*?"%'#<>|=]/g, '-').replace(/[^a-zA-Z0-9]*$/g, '');
    // get the extension
    var e = n.substring(n.lastIndexOf('.') + 1);
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

    var file = {
      FileName: newName,
      Data: data,
      Length: data.byteLength,
      ListItemAllFields: { Title: t }
    };

    if (desc && this.file.notes)
      file.ListItemAllFields[this.file.notes] = desc;

    results.push(file);

    if (!this.file.extract)
      return;

    if (fileName.toLowerCase().endsWith(".zip"))
      await this.zips(data, results);
    if (fileName.toLowerCase().endsWith(".msg"))
      await this.msgs(data, results);
    if (fileName.toLowerCase().endsWith(".eml"))
      await this.emls(data, results);

    this.office.loading = false;
    this.chRef.detectChanges();
  }

  // extract zip files and append to results
  async zips(data: ArrayBuffer, results: Array<any>) {
    try {
      var zip = await loadAsync(data);
      var files = Object.keys(zip.files);
      files.forEach(async (file) => {
        try {
          var buffer: ArrayBuffer | undefined = await zip.file(file)?.async('arraybuffer');
          if (buffer) {
            await this.appendFile(file, buffer, results, `Date: ${zip.files[file].date}`);
          }
        } catch (e) { }
      });
    } catch (e) {
      // zip is uploaded so any extracted elements are only nice to have
    }
  }

  // extract and append msg email attachments to results
  async msgs(data: ArrayBuffer, results: Array<any>) {
    try {
      // new MsgReader(data) doesnt seem to work and .default is not recognised but ['default'] works somehow
      var msgReader = new MsgReader['default'](data) as MsgReader;
      // triggered the parser
      var fileData = msgReader.getFileData();
      // if no sender name then its not an email
      if (!('senderName' in fileData))
        return;
      // get the email date
      var h = fileData.headers?.split('\n').filter((x: string) => x.startsWith('Date: '));
      var received = h && h.length > 0 ? new Date(h[0].replace('Date: ', '')) : new Date();
      // get all attachments
      fileData.attachments?.forEach(async (attachment) => {
        if (attachment.attachmentHidden)
          return;
        try {
          var file = msgReader.getAttachment(attachment);
          await this.appendFile(file.fileName, file.content.buffer as ArrayBuffer, results, `Sent: ${received}`);
        } catch (e) { }
      });
    } catch (e) {
      // msg is uploaded so any extracted elements are only nice to have
    }
  }

  // extract and append eml email attachments to results
  async emls(data: ArrayBuffer, results: Array<any>) {
    try {
      // reads the email string data into a json object
      readEml(new TextDecoder().decode(data), (err, ReadEmlJson) => {
        if (err || !ReadEmlJson || !ReadEmlJson.attachments)
          return;
        var received = typeof ReadEmlJson.date == "string" ? new Date(ReadEmlJson.date) : ReadEmlJson.date || new Date();
        ReadEmlJson.attachments.forEach(async (attachment: any) => {
          if (attachment.inline)
            return;
          try {
            // work out the name from id which is more consistent across sources, otherwise from name. 
            var name = attachment.id?.replace(/^</, '').replace(/>$/, '').split('@')[0];
            if (!name || name.indexOf('.') < 0)
              name = attachment.name;
            await this.appendFile(name, Uint8Array.from(atob(attachment.data64), c => c.charCodeAt(0)).buffer, results, `Sent: ${received}`);
          } catch (e) { }
        });
      });
    } catch (e) {
      // eml is uploaded so any extracted elements are only nice to have
    }
  }

  // dragging and dropping, hover
  over(evt: any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.filesOver = true;
  }

  // dragging and dropping, unhover
  leave(evt: any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.filesOver = false;
  }

  // dragging and dropping, drop
  async drop(evt: any) {
    evt.preventDefault();
    evt.stopPropagation();

    if (evt.dataTransfer?.files?.length > 0)
      await this.add(evt.dataTransfer);
    if (evt.dataTransfer?.items?.length > 0)
      await this.outlook(evt.dataTransfer);

    // after files then drop the shake
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

  // select user
  async selectedUser(res: any): Promise<void> {
    if (!this.spec['odata.context'])
      return;

    // ensure correct schema
    if (this.get('TypeAsString') == 'UserMulti' && (!this.form[this.field + 'Id'] || !this.form[this.field + 'Id'].__metadata))
      this.form[this.field + 'Id'] = {
        __metadata: { type: "Collection(Edm.Int32)" },
        results: !this.form[this.field + 'Id'] ? [] : this.form[this.field + 'Id'].results ? this.form[this.field + 'Id'].results : typeof this.form[this.field + 'Id'] == "object" ? this.form[this.field + 'Id'] : [this.form[this.field + 'Id']]
      }
    if (this.get('TypeAsString') == 'User' && this.form[this.field + 'Id'] && this.form[this.field + 'Id'].results)
      this.form[this.field + 'Id'] = this.form[this.field + 'Id'].results.length > 0 ? this.form[this.field + 'Id'].results[0] : null;

    // use click item
    if (res) {
      // already selected, do nothing
      if (this.get('TypeAsString') == 'UserMulti' && this.display.filter((x: any) => {
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
      await this.onchange(this);

    this.chRef.detectChanges();
  }

  // load list data only has IDs so expand the object
  displayUser(user: any): string {
    if (!this.spec['odata.context'] || !user)
      return user || '';

    var u = this.display.filter((x: any) => {
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
      this.spec['odata.context'].web.getUserById(user)().then((u: any) => {
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

        this.chRef.detectChanges();
      });
    }

    return '';
  }

  // removes a user
  async removeUser(usr: any): Promise<void> {
    if (!usr) {
      this.form[this.field + 'Id'] = null;
      this.display = [];
      return;
    }

    this.form[this.field + 'Id'].results.splice(this.form[this.field + 'Id'].results.indexOf(usr), 1);
    this.display = this.display.filter((x: any) => {
      return x.Id != usr
    })

    if (typeof this.onchange == "function")
      await this.onchange(this);

    this.chRef.detectChanges();
  }

  // trigger user search
  onUpUser(key): void {
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

  async onUpUserSearch(text: string): Promise<void> {
    if (!this.spec['odata.context'])
      return;

    var url = (await this.spec['odata.context'].web()).ServerRelativeUrl;

    // set the user partial being searched
    this.UserQuery.queryParams.QueryString = this.name;
    // set group each query to adapt to changes
    this.UserQuery.queryParams.SharePointGroupID = parseInt(this.get('SelectionGroup') || 0);
    // ensure up to date digest for http posting
    var token: any = await fetch(url + '/_api/contextinfo', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json;odata=verbose'
      }
    });
    let digest = await token.json();
    // query users api, no pnp endpoint for this
    var search: any = await fetch(url + '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': digest.d.GetContextWebInformation.FormDigestValue
      },
      body: JSON.stringify(this.UserQuery)
    });
    let res: any = await search.json();
    this.users = [];
    let allUsers = JSON.parse(res.d.ClientPeoplePickerSearchUser);
    allUsers.filter(x => {
      return x.EntityData.Email && !~x.Key.indexOf('_adm') && !~x.Key.indexOf('adm_')
    }).forEach(user => {
      this.users = [...this.users, user];
    });

    this.chRef.detectChanges();
  }
}


