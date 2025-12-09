import { Component, OnInit, OnDestroy, Input, ElementRef, ChangeDetectorRef, ErrorHandler, Output, EventEmitter } from '@angular/core';
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import { Editor, NgxEditorModule, Toolbar } from 'ngx-editor';
import MsgReader from '@kenjiuno/msgreader';
import { Attachment, readEml } from 'eml-parse-js';
import { loadAsync } from 'jszip';
import { Subject } from 'rxjs';
import { debounceTime, distinctUntilChanged } from 'rxjs/operators';
import { SharepointChoiceUtils } from './sharepoint-choice.utils';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SharepointChoiceLogging } from './sharepoint-choice.logging';
import { SharepointChoiceForm, SharepointChoiceList, SharepointChoiceField, SharepointChoiceAttachment, SharepointChoiceUser } from './sharepoint-choice.models';

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
  @Input() prefix: string = ''; // prefix name attributes for uniqness, usefull for nesting

  @Input() form!: SharepointChoiceForm; // form containing field, varies based on list therefore uses flexible interface

  @Input() field!: string; // internal field name on form object, used for push back and against spec

  @Input() spec!: SharepointChoiceList; // spec of field loaded from list
  @Input() versions?: SharepointChoiceForm[]; // version history of this field to display if presented, varies based on list therefore object not defined explicitly
  @Input() override?: string | SharepointChoiceField; // manually override any spec above. prefer send as string as passing object kills large form performance

  // could also input hidden rather than elRef but that doesnt pick up non angular html hidden or parent hidden
  @Input() disabled: boolean = false; // get disabled state from outside
  @Output() change = new EventEmitter<{ field: string, value: any }>(); // emit changes to parent through (change) binding

  @Input() text?: { // override text for field
    pattern?: string, // regex pattern for validation
    height?: number, // height of text area in px

    // should move from call backs that depend on parent being passed in to @Output keys sent @Input search results @Output selected but it would place these for all field types
    search?: Function, // search via api for drop down options
    select?: Function, // upon selection in drop down call back function
    parent?: any // parent object that the control belongs to for call backs or specific search functions
  };

  @Input() select?: { // override select for field
    none?: string, // none option text instead of null
    other?: string, // Other fill-in option text, will override to allow other

    filter?: Function // filter choices by a function
  };

  @Input() file?: { // override file for field
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
    spec?: SharepointChoiceList // field spec for additional fields, 
  };

  declare editor?: Editor;
  declare toolbar: Toolbar;
  declare tooltip?: boolean;
  declare filesOver?: boolean;
  declare name?: string;
  declare loading?: Array<number>;
  declare versionsDisplayed?: boolean;

  declare users: SharepointChoiceUser[];

  declare filterMulti?: string;
  declare unused: string;
  declare results: Array<object>; // varies based on the search source therefore not explicitly defined
  declare pos: number;
  declare office: {
    type: string | null,
    loading: boolean
  };

  declare sort?: string;
  declare filter: string;

  public textKey?: Subject<string>;
  public userKey?: Subject<string>;

  private overridePrevious?: string;
  private overrideParsed?: object;

  private display: SharepointChoiceUser[];

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
    if (!this.disabled)
      this.disabled = false;

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
  }

  // on init, destroy
  ngOnInit(): void {
  }
  ngOnDestroy(): void {
    this.editor?.destroy();
    this.textKey?.complete();
    this.userKey?.complete();
  }

  /* 
  Common parts between multiple field types or minimal functions
  */

  // show or hide tooltips 
  showHideTooltip(show: boolean): void {
    this.tooltip = show;
  }

  // are there different field version values shown
  versionsToggle(): string {
    if (!this.versionsDisplayed)
      this.versionsDisplayed = true;
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
    let p = parseFloat(e.replace(/[^0-9\.]/g, ''));
    if (isNaN(p)) {
      this.form[this.field] = null
      return;
    }
    let min = this.get('Min');
    if (min != null && p < min)
      p = min;
    let max = this.get('Max');
    if (max != null && p > max)
      p = max;
    this.form[this.field] = p;

    this.changed();
  }

  dateSet(e: Date|string|null): void {
    this.form[this.field] = e;

    this.changed();
  }

  // field required based on spec but required is not needed for hidden/disabled items
  required(): boolean {
    if (this.disabled || this.elRef.nativeElement.hidden || !this.get('Required'))
      return false;
    return true;
  }

  // get outcomes of non standard fields into a plain text field for [required] to be triggered automatically 
  validator(type?: string): string {
    switch (type) {
      case 'User':
        return this.form[this.field + 'Id'] ? 'true' : '';
      case 'UserMulti':
        return (this.form[this.field + 'Id'] && this.form[this.field + 'Id'].results && this.form[this.field + 'Id'].results.length > 0) ? 'true' : '';
      case 'Attachments':
        return this.attachments().length > 0 ? 'true' : '';
      default:
        return this.form[this.field] ? 'true' : '';
    }
  }

  // max length character countdown
  remaining(max?: number): number {
    let m = this.get('MaxLength');
    // no max limit in field spec or because of field type then use a number bigger than any remaining() >= ... values
    if (!m && !max)
      return 255;
    return (m ?? max) - (this.form[this.field] ?? '').length;
  }

  // gets the required field properties and/or any overrides to determine which field type etc to display
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
    let p: any = this.overrideParsed ? this.overrideParsed[t] : null;
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
        if (!this.userKey) {
          this.userKey = new Subject<string>();
          this.userKey.pipe(
            debounceTime(250),
            distinctUntilChanged()
          ).subscribe((key) => this.onUpUserSearch(key));
        }
      } else if (p == 'User') {
        if (!this.userKey) {
          this.userKey = new Subject<string>();
          this.userKey.pipe(
            debounceTime(250),
            distinctUntilChanged()
          ).subscribe((key) => this.onUpUserSearch(key));
        }
        // if its a url, ensure the correct object type and clone data into url for flat stored occurrences 
      } else if (p == 'URL') {
        if (!this.form[this.field] || this.form[this.field].Description === undefined || this.form[this.field].Description === null)
          this.form[this.field] = {
            Description: '',
            Url: this.form[this.field] || ''
          }
        // if its attachments, ensure the object is the correct type and office fired if needed
      } else if (p == 'Attachments') {
        this.officeAddin();
        if (!this.form[this.field] || !this.form[this.field].results)
          this.form[this.field] = {
            results: []
          }
      } else if (p == 'Text') {
        if (!this.textKey && this.text?.search) {
          this.textKey = new Subject<string>();
          this.textKey.pipe(
            debounceTime(250),
            distinctUntilChanged()
          ).subscribe((key) => this.onUpTextSearch(key));
        }
      }

      // if the field is empty (not set null) and its not a person field (+Id) then set the default value only if there is one
      if (this.form[this.field] === undefined && !this.form[this.field + 'Id']) {
        let d = this.get('DefaultValue');
        if (d && p == 'MultiChoice')
          this.form[this.field].results = [d];
        else if (d && p == 'UserMulti')
          this.form[this.field + "Id"].results = [d];
        else if (d && p == 'User')
          this.form[this.field + "Id"] = d;
        else if (d)
          this.form[this.field] = (
            p == "DateTime" ? new Date(d) :
              p == "Integer" ? parseInt(d) :
                p == "Number" ? parseFloat(d) :
                  p == "Boolean" ? (d == '1' ? true : false) :
                    d
          );
      }

      // always disable read only fields initially, may get overridden later but thats the consumers responsibility
      let r = this.get('ReadOnlyField');
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

    return p?.results ?? p;
  }

  // any field changes trigger for relevant updates
  changed(trigger: boolean = false): void {
    if (trigger)
      this.chRef.detectChanges();
    this.change.emit({ field: this.field, value: this.form[this.field]?.results ?? this.form[this.field] });
  }

  // on multi select, not using ctrl key
  changedM(e: Event): void {
    e.stopPropagation();
    var target = (e.target as HTMLElement);
    let scrollTop = 0;
    if (target.parentElement)
      scrollTop = target.parentElement.scrollTop;

    let v = target.innerText;
    // ensure object type is correct
    if (!this.form[this.field] || !this.form[this.field].__metadata)
      this.form[this.field] = {
        __metadata: { type: "Collection(Edm.String)" },
        results: !this.form[this.field] ? [] : this.form[this.field].results ? this.form[this.field].results : typeof this.form[this.field] == "object" ? this.form[this.field] : [this.form[this.field]]
      }
    // if there are selected results set the field to add/remove the most recent click
    let i = this.form[this.field].results.indexOf(v);
    if (i >= 0)
      this.form[this.field].results.splice(i, 1);
    else
      this.form[this.field].results.push(v);

    const tmp = this.form[this.field].results;
    this.form[this.field].results = [];
    for (let i = 0; i < tmp.length; i++)
      this.form[this.field].results[i] = tmp[i];

    if (target.parentElement) {
      setTimeout((function () { target.parentElement ? target.parentElement.scrollTop = scrollTop : null; }), 10);
      setTimeout((function () { target.parentElement ? target.parentElement.focus() : null; }), 10);
    }

    this.changed();
  }

  // append only changes needs 1 way bind to form
  changedA(e: Event): void {
    this.form[this.field] = e.target ? (e.target as HTMLInputElement).value : e;

    this.changed();
  }

  /* 
  Common parts between text fields
  */

  async onUpText(key:string|undefined): Promise<void> {
    if (!this.textKey)
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
    if (!this.text?.search)
      return;

    this.results = await this.text.search(this.form[this.field], this.text.parent);

    this.pos = -1;
    this.chRef.detectChanges();
  }

  async selectedText(res: SharepointChoiceForm|null): Promise<void> {
    if (!this.text?.search)
      return;

    if (!res) {
      this.form[this.field] = null;
    } else {
      this.form[this.field] = res[this.field];
      if (this.text.select)
        await this.text.select(res, this.text.parent);
    }

    this.results = [];
    this.changed(true);
  }

  /* 
  Common parts between choice fields
  */

  // many choices render bigger box
  multiLargeorSmall(): string {
    return this.choices().length > 10 ? 'multilarge' : 'multismall'
  }

  // choices need filtering
  choices(): string[] {
    // get choices from list
    let choices = this.get('Choices');
    // use any provided filter, pass in the form for context or specific filters but not the remainder of this object
    if (typeof this.select?.filter == "function")
      choices = choices.filter((c: string, i: number, a: string[]) => this.select?.filter ? this.select.filter(c, i, a, { form: this.form }) : true);
    // common filters
    return choices.filter((x: string) => {
      // exclude unselected items on disabled fields
      if (this.disabled && this.form[this.field] && this.form[this.field].results && !this.form[this.field].results.includes(x))
        return false;
      // filter on search above multichoice field
      if (this.filterMulti && !x.toLowerCase().includes(this.filterMulti.toLowerCase()))
        return false;
      // else true
      return true;
    })
  }

  // selected field option not in available choices, i.e. other
  inChoices(c: string[]): boolean {
    if (!this.form[this.field])
      return true;
    return c.indexOf(this.form[this.field]) >= 0;
  }

  // on single selection change
  selChangeS(v: string): void {
    this.form[this.field] = v;

    this.changed();
  }

  /* 
  Common parts between file upload field types
  */

  // files post filtering
  attachments(): SharepointChoiceAttachment[] {
    if (!this.form[this.field] || !this.form[this.field].results)
      return [];
    let v = this.file?.view || 0;
    return this.form[this.field].results
      .filter((f: SharepointChoiceAttachment) => {
        if (v == 0 || !this.file?.archive || f.ListItemAllFields == null)
          return true;
        if (v == 1 && !f.ListItemAllFields[this.file.archive])
          return true;
        if (v == -1 && f.ListItemAllFields[this.file.archive])
          return true;
        return false;
      })
      .filter((f: SharepointChoiceAttachment) => {
        if (!this.filter || !this.file?.doctype)
          return true;
        if (!f.ListItemAllFields || !f.ListItemAllFields[this.file.doctype])
          return true;
        if (f.ListItemAllFields[this.file.doctype] == this.filter)
          return true;
        return false;
      })
      .sort((a: SharepointChoiceAttachment, b: SharepointChoiceAttachment) => {
        if (!this.sort || this.sort == '-')
          this.sort = 'Created';
        let s = this.sort.replace(/^[-\+]/, '');
        let o = this.sort.startsWith('-') ? -1 : 1;
        if (!a.ListItemAllFields || !a.ListItemAllFields[s])
          return -2;
        if (!b.ListItemAllFields || !b.ListItemAllFields[s])
          return 2;
        if (a.ListItemAllFields[s] < b.ListItemAllFields[s])
          return o;
        if (a.ListItemAllFields[s] > b.ListItemAllFields[s])
          return -o;
        return 0;
      });
  }

  friendlyName(name?: string): string {
    if (!name)
      return '';
    return name.replace(/_x0020_/g, ' ').replace(/([a-z])([A-Z])/g, '$1 $2');
  }

  hasChecked(): boolean {
    if (!this.form[this.field] || !this.form[this.field].results)
      return false;
    // check if there are any files that have been checked
    return this.form[this.field].results.some((f: SharepointChoiceForm) => f['Checked']);
  }

  setClasses(e: Event): void {
    let t = (e.target as HTMLSelectElement);
    // set the class of the file
    this.form[this.field].results.forEach((f: SharepointChoiceAttachment) => {
      if (f.Checked) {
        if (f.ListItemAllFields)
          f.ListItemAllFields['DocumentType'] = t?.value;
        f.Checked = false;
      }
    });
    if (t)
      t.value = '';

    this.changed();
  }

  changeSort(field?: string) {
    if (!field)
      return;
    if (this.sort == '-' + field) {
      this.sort = '+' + field;
    } else if (this.sort == '+' + field) {
      this.sort = '';
    } else {
      this.sort = '-' + field;
    }
  }

  additionalKeys(o?: SharepointChoiceList): Array<string> {
    // get keys of object
    if (!o || typeof o != "object" || this.field == 'Attachments')
      return [];
    return Object.keys(o).filter(k => k != this.file?.doctype && k != this.file?.notes);
  }

  loadOrGenerateSpec(field?: string, type?: string): SharepointChoiceList {
    let s: SharepointChoiceList = {};
    if (!field)
      return s;
    if (this.file?.spec && this.file.spec[field])
      s[field] = this.file.spec[field];
    else
      s[field] = {
        TypeAsString: (type || 'Text') as any,
        InternalName: field,
        Title: '',
        Choices: this.file?.doctypes || []
      };
    return s;
  }

  getFieldSpec(key: string): SharepointChoiceField | undefined {
    const value = this.file?.spec?.[key];
    return value && typeof value === 'object' && 'web' in value === false ? value : undefined;
  }

  usedTypes(): Array<string> {
    // get initial types
    let types = this.file?.doctypes || [];
    // get all types used in the attachments
    if (this.file?.doctype && this.form[this.field] && this.form[this.field].results)
      types = this.form[this.field].results
        .map((f: SharepointChoiceAttachment) => f.ListItemAllFields && this.file?.doctype && this.file.doctype in f.ListItemAllFields ? f.ListItemAllFields[this.file.doctype] : null)
        .filter((f: string) => f)
        .sort();
    // remove duplicates
    types = [...new Set(types)];
    return types;
  }

  newTab(f: SharepointChoiceAttachment, e: Event): void {
    window.open(`${f.ServerRelativeUrl || '#'}?Web=1`, '_blank');
  }

  width(): string {
    let c = 0;
    if (this.file?.doctypes)
      for (let i = 0; i < this.file.doctypes.length; i++) {
        let l = this.file.doctypes[i].length
        if (l > c)
          c = l;
      }
    if (c == 0)
      return '';
    return `width: ${c}ch`;
  }

  async delete(f?: SharepointChoiceAttachment, a: boolean = false) {
    if (!f)
      return;
    if (!f.ServerRelativeUrl) {
      // not uploaded then exclude from potential upload
      this.form[this.field].results = this.form[this.field].results.filter((x: SharepointChoiceAttachment) => x.FileName != f.FileName);
    } else {
      // if uploaded already
      if (!this.file?.archive || this.field == 'Attachments') {
        // no archive flag or its attachments so no archiving, then flag for deletion
        if (a && !window.confirm('Are you sure you wish to delete this file?'))
          return;
        f.Deleted = a;
      } else {
        // toggle archived or delete flag
        if (f.ListItemAllFields?.[this.file.archive] && (a || f.Deleted)) {
          if (a && !window.confirm('Are you sure you wish to delete this file?'))
            return;
          f.Deleted = a;
        } else if (f.ListItemAllFields)
          f.ListItemAllFields[this.file.archive] = a;
      }
    }

    this.changed(true);
  }

  // add attachment to array
  async add(file: HTMLInputElement|DataTransfer) {
    if (!file.files || file.files.length == 0)
      return;
    // read the files into the files array
    let files: Array<File> = [];
    for (let i = file.files.length - 1; i >= 0; i--)
      files.push(file.files[i]);
    // copy these outside for reuse in the loop
    let ths = this;
    let remaining = files.length;
    // loop the array in forEach for variable isolation
    files.forEach((f: File) => {
      let reader = new FileReader();
      reader.onload = async function (event: ProgressEvent<FileReader>) {
        try {
          await ths.appendFile(f.name, event.target?.result as ArrayBuffer, ths.form[ths.field].results);

          remaining--;

          // if last file added then clear the file input
          if (remaining == 0 && file instanceof HTMLInputElement)
            setTimeout(() => {
              file.value = '';
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
  async outlook(transfer: DataTransfer): Promise<void> {
    let spc = new SharepointChoiceUtils();

    function mailType(transfer: DataTransfer, type: string) {
      let item = transfer.getData(type);
      if (!item)
        return null;
      return JSON.parse(item)
    }

    // one or more email messages dropped
    let maillistrow = mailType(transfer, 'multimaillistmessagerows') || mailType(transfer, 'maillistrow');
    if (maillistrow) {
      let errors: Array<string> = [];
      for (let i = 0; i < maillistrow.mailboxInfos.length; i++) {
        let fileName = maillistrow.subjects[i].trim() + ".eml";

        try {
          // must double url encode any / in the message id
          let fileContent: string = await spc.callApi(
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
        var getAttachment: {lastModifiedDateTime: string, contentBytes: string} = await spc.callApi(
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
    if (document.getElementById('officejs') || (window as any).Office)
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
    let s = document.createElement('script');
    s.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
    s.id = 'officejs';
    document.head.appendChild(s);
    // capture the office type for later use along with triggering change detection if needed
    let ths = this;
    s.addEventListener('load', () => {
      // only once office.js loaded
      if ('Office' in window) {
        let office: typeof Office = (window as any).Office;
        if (!office)
          return;
        // on ready should be ran soon after office.js is loaded trigger
        office.onReady(info => {
          // capture what office type of addin for later use
          if (ths.office.type != info.host.toString()) {
            ths.office.type = info.host.toString();
            ths.chRef.detectChanges();
          }
        });
      }
    }, false);
  }

  // import from office addin selection or document panel
  importOutlook() {
    this.office.loading = true;
    let office: typeof Office = (window as any).Office;
    if (!office)
      return;

    // if the adding type is outlook then get the selected email(s)
    var spc = new SharepointChoiceUtils();
    var ths = this;

    if (office.context.mailbox['initialData']?.isFromSharedFolder) {
      office.context.mailbox.item?.getSharedPropertiesAsync(shared => {
        if (shared.status === Office.AsyncResultStatus.Failed)
          return;
        this.getMailItem(office, shared?.value?.targetMailbox, spc, ths);
      });
    } else {
      this.getMailItem(office, null, spc, ths);
    }
  }

  getMailItem(office: typeof Office, targetMailbox: string | null, spc: SharepointChoiceUtils, ths: SharepointChoiceComponent) {
    office.context.mailbox.getSelectedItemsAsync(async (asyncResult: Office.AsyncResult<Office.SelectedItemDetails[]>) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed)
        return;

      var errors: Array<string> = [];
      for (var m in asyncResult.value) {
        var message = asyncResult.value[m];
        try {

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
    let office: typeof Office = (window as any).Office;
    if (!office)
      return;

    // if the adding type is word or excel then get the current document
    let docDataSlices: Office.FileType.Text[] | Office.FileType.Compressed[] = [];
    let slicesReceived = 0;
    let file: Office.File;
    let ths = this;

    try {
      // get the file in 64k slices until complete
      office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (asyncResult: Office.AsyncResult<Office.File>) => {
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
      function processSlice(sliceResult: Office.AsyncResult<Office.Slice>) {
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
        let docData:(Office.FileType.Text | Office.FileType.Compressed)[] = [];
        for (let i = 0; i < docDataSlices.length; i++)
          docData = docData.concat(docDataSlices[i]);

        // convert char codes to string
        let fileContent = new String();
        for (let j = 0; j < docData.length; j++)
          fileContent += String.fromCharCode(docData[j]);

        // try to get a file name from the document
        var fileName = office.context.document.url?.split('/')?.pop()?.split('\\')?.pop();
        if (!fileName) {
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

  async appendFile(fileName: string, data: ArrayBuffer, results: SharepointChoiceAttachment[], desc?: string) {
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

    var file:SharepointChoiceAttachment = {
      FileName: newName,
      Data: data,
      Length: data.byteLength,
      ListItemAllFields: { Title: t }
    };

    if (desc && this.file?.notes && file.ListItemAllFields)
      file.ListItemAllFields[this.file.notes] = desc;

    results.push(file);

    if (!this.file?.extract)
      return;

    if (fileName.toLowerCase().endsWith(".zip"))
      await this.zips(data, results);
    if (fileName.toLowerCase().endsWith(".msg"))
      await this.msgs(data, results);
    if (fileName.toLowerCase().endsWith(".eml"))
      await this.emls(data, results);

    this.office.loading = false;
    this.changed(true);
  }

  // extract zip files and append to results
  async zips(data: ArrayBuffer, results: SharepointChoiceAttachment[]) {
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
  async msgs(data: ArrayBuffer, results: SharepointChoiceAttachment[]) {
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
  async emls(data: ArrayBuffer, results: SharepointChoiceAttachment[]) {
    try {
      // reads the email string data into a json object
      readEml(new TextDecoder().decode(data), (err, ReadEmlJson) => {
        if (err || !ReadEmlJson || !ReadEmlJson.attachments)
          return;
        var received = (typeof ReadEmlJson.date == "string" ? new Date(ReadEmlJson.date) : ReadEmlJson.date) || new Date();
        ReadEmlJson.attachments.forEach(async (attachment: Attachment) => {
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
  over(evt: DragEvent) {
    if (this.disabled)
      return;
    evt.preventDefault();
    evt.stopPropagation();
    this.filesOver = true;
  }

  // dragging and dropping, unhover
  leave(evt: DragEvent) {
    if (this.disabled)
      return;
    evt.preventDefault();
    evt.stopPropagation();
    this.filesOver = false;
  }

  // dragging and dropping, drop
  async drop(evt: DragEvent) {
    if (!evt)
      return;
    if (this.disabled)
      return;
    evt.preventDefault();
    evt.stopPropagation();

    if (evt.dataTransfer?.files?.length && evt.dataTransfer?.files?.length > 0)
      await this.add(evt.dataTransfer);
    if (evt.dataTransfer?.items?.length && evt.dataTransfer?.items?.length > 0)
      await this.outlook(evt.dataTransfer);

    // after files then drop the shake
    this.filesOver = false;

    // if no transfer on drop and not chromium based add the meta tag to allow the drop next time
    if (!evt.dataTransfer && !(window as any).chrome) {
      let m = document.createElement("meta");
      m.httpEquiv = "X-UA-Compatible";
      m.content = "chrome=1";
      document.head.appendChild(m);
    }
  }

  /* 
  Common parts between user field types
  */

  // select user
  async selectedUser(res: SharepointChoiceUser|null): Promise<void> {
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
      if (this.get('TypeAsString') == 'UserMulti' && this.display.filter((x: SharepointChoiceUser) => {
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
          Id: u.Id,
          Title: res.Title ?? res.DisplayText,
          LoginName: res.LoginName ?? res.Key,
        });
      } else {
        this.form[this.field + 'Id'] = u.Id;
        this.display = [{
          DisplayText: res.DisplayText,
          Key: res.Key,
          Id: u.Id,
          Title: res.Title ?? res.DisplayText,
          LoginName: res.LoginName ?? res.Key,
        }];
      }
    }

    // clear search fields
    this.name = '';
    this.users = [];

    this.changed(true);
  }

  // load list data only has IDs so expand the object
  displayUser(user?: number): string {
    if (!this.spec['odata.context'] || !user)
      return user?.toString() || '';

    var u = this.display.filter((x: SharepointChoiceUser) => {
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
      this.spec['odata.context'].web.getUserById(user)().then((u: SharepointChoiceUser) => {
        // update the display table
        this.display.push({
          DisplayText: u.Title,
          Key: u.LoginName,
          Id: u.Id,
          Title: u.Title ?? u.DisplayText,
          LoginName: u.LoginName ?? u.Key,
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
  removeUser(usr: number|null): void {
    if (!usr) {
      this.form[this.field + 'Id'] = null;
      this.display = [];
      return;
    }

    this.form[this.field + 'Id'].results.splice(this.form[this.field + 'Id'].results.indexOf(usr), 1);
    this.display = this.display.filter((x: SharepointChoiceUser) => {
      return x.Id != usr
    });

    this.changed();
  }

  // trigger user search
  onUpUser(key:string|undefined): void {
    if (!this.userKey)
      return;

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
    if (!this.name)
      return;

    var url = (await this.spec['odata.context'].web()).ServerRelativeUrl;

    // ensure up to date digest for http posting
    var token: Response = await fetch(url + '/_api/contextinfo', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json;odata=verbose'
      }
    });
    let digest = await token.json();
    // query users api, no pnp endpoint for this
    var search: Response = await fetch(url + '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': digest.d.GetContextWebInformation.FormDigestValue
      },
      body: JSON.stringify({
        queryParams: {
          // set the user partial being searched
          QueryString: this.name,
          MaximumEntitySuggestions: 10,
          AllowEmailAddresses: true,
          AllowOnlyEmailAddresses: false,
          PrincipalSource: 15,
          PrincipalType: 1,
          // set group each query to adapt to changes
          SharePointGroupID: parseInt(this.get('SelectionGroup') || '0')
        }
      })
    });
    let res = await search.json();
    this.users = [];
    let allUsers = JSON.parse(res.d.ClientPeoplePickerSearchUser);
    allUsers.filter((x: SharepointChoiceUser) => {
      return x.EntityData?.Email && !x.Key.includes('_adm') && !x.Key.includes('adm_')
    }).forEach((user: SharepointChoiceUser) => {
      this.users = [...this.users, user];
    });

    this.chRef.detectChanges();
  }
}