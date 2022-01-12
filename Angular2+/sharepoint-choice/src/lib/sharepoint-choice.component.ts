import { Component, OnInit, Input, ElementRef, ViewEncapsulation } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { catchError, tap, mergeMap } from 'rxjs/operators';
import { Observable } from 'rxjs';
import { UserQuery, User } from "./Models";
import pnp from '@pnp/pnpjs';
import { Logger, LogLevel } from "@pnp/logging";
import { PnPLogging } from './PnPLogging';
import { App } from './App'

@Component({
  selector: 'app-choice',
  templateUrl: './sharepoint-choice.component.html',
  styleUrls: ['./sharepoint-choice.component.scss'],
  encapsulation: ViewEncapsulation.Emulated
})
export class SharepointChoiceComponent implements OnInit {
  @Input() form!: any[string]; // form containing field
  @Input() spec!: any[any[string]]; // spec of field loaded from list
  @Input() field!: string; // internal field name on form object, used for push back and against spec
  @Input() override!: string; // override any spec above. sent as string as passing object kills large form performance

  @Input() none!: string; // none option text instead of null
  @Input() other!: string; // Other fill-in option text, will override to allow other
  @Input() filter!: Function; // filter choices by a function

  @Input() disabled!: boolean;

  constructor(
    private elRef: ElementRef,
    private _http: HttpClient
  ) { }

  declare tinymceOptions: object;
  declare tooltip: boolean;
  declare filesOver: boolean;
  declare name: string;
  declare users: User[];
  declare display: any;
  declare loading:any[number];
  declare key: string;

  // schema setup
  UserQuery: UserQuery = {
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

  // on init
  ngOnInit(): void {
    pnp.sp.setup({sp:{baseUrl:this.spec['odata.metadata']}});
    
    Logger.subscribe(new PnPLogging());
    Logger.activeLogLevel = LogLevel.Warning;

    // rich text field
    this.tinymceOptions = {
        resize: false,
        height: 500,
        menubar: false,
        plugins: "textcolor lists table link paste",
        toolbar: "forecolor | bold italic underline | bullist numlist outdent indent | table | link",
        statusbar: false,
        debounce: false,
        paste_data_images: true
    };

    // user(s)
    this.users = [];
    this.display = [];
    this.key = App.TinyMCEKey;
  }

  people(): string {
    return this.form[this.field + 'Id'] && (!this.form[this.field + 'Id'].results || this.form[this.field + 'Id'].results.length > 0) ? 'true' : null;
  }

  attach(): string {
    return this.attachments().length > 0 ? 'true' : null;
  }

  // select user
  async selectedUser(res:any): Promise<void> {
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
      var u = await pnp.sp.web.ensureUser(res.Key);

      // add to field
      if (this.get('TypeAsString') == 'UserMulti') {
        this.form[this.field + 'Id'].results.push(u.data.Id);
        this.display.push({
          DisplayText: res.DisplayText,
          Key: res.Key,
          Id: u.data.Id
        });
      } else {
        this.form[this.field + 'Id'] = u.data.Id;
        this.display = [{
          DisplayText: res.DisplayText,
          Key: res.Key,
          Id: u.data.Id
        }];
      }
    }

    // clear search fields
    this.name = '';
    this.users = [];
  }

  // load list data only has IDs so expand the object
  displayUser(user:any): string {
    if (!user)
      return '';
    var u = this.display.filter((x:any) => {
      return x.Id == user
    });
    if (u.length > 0 && u[0].DisplayText)
      return u[0].DisplayText;

    // setup context late to adapt to changes
    if (!this.loading)
      this.loading = [];
    if (this.loading.indexOf(user) < 0) {
      this.loading.push(user);
      var ths = this;
      pnp.sp.web.getUserById(user).get().then(u => {
        ths.display.push({
          DisplayText: u.Title,
          Key: u.LoginName,
          Id: u.Id
        });
      });
    }
    return '';
  }

  // removes an user
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
  }

  // trigger user search
  onUp(): void {
    if (this.name.length < 3) {
      this.users = [];
      return;
    }

    this.UserQuery.queryParams.QueryString = this.name;
    // set group each query to adapt to changes
    this.UserQuery.queryParams.SharePointGroupID = parseInt(this.get('SelectionGroup') || 0);
    this.getResponse(this.UserQuery)
      .subscribe(
        (res) => {
          this.users = [];
          const allUsers: User[] = JSON.parse(res.d.ClientPeoplePickerSearchUser);
          allUsers.filter(x => {
            return x.EntityData.Email && !~x.Key.indexOf('_adm') && !~x.Key.indexOf('adm_')
          }).forEach(user => {
            this.users = [...this.users, user];
          });
        }, (error) => {
          Logger.log({
            message: `Inside - getResponse(${error})`,
            level: LogLevel.Error
          });
        });
  }

  // get users via service
  getResponse(query: UserQuery): Observable<any> {
    return this._http.post(
      this.get('Context') + '/_api/contextinfo',
      '',
      {
        headers: new HttpHeaders({
          'Content-Type': 'application/json',
          'Accept': 'application/json;odata=verbose'
        })
      }
    ).pipe(
      mergeMap((xRequest: any) => {
        return this._http.post(
          this.get('Context') + '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser',
          query,
          {
            headers: new HttpHeaders({
              accept: 'application/json;odata=verbose',
              'X-RequestDigest': xRequest.d.GetContextWebInformation.FormDigestValue
            })
          }
        ).pipe(
          tap(httpres => {
            console.log('Fetched Data')
          }),
          catchError(error => {
              throw error.error.error.message.value;
          })
        );
      })
    );
  }
  
  // max length countdown
  remaining(): number {
    var m = this.get('MaxLength');
    if (!m && this.get('TypeAsString') == 'Text')
      m = 255;
    if (!m)
      return 15;
    return m - (this.form[this.field] || '').length;
  }

  // files post filtering
  attachments(): any[string] {
    if (!this.form[this.field] || !this.form[this.field].results)
      return [];
    this.form[this.field].results.forEach((r:any) => {
      r.Prefix = r.Prefix || (~r.FileName.indexOf('-') ? r.FileName.split('-')[0] : '');
    });
    var prefix:any = this.get('Prefix') || '';
    return this.form[this.field].results.filter((r:any[string]) => {
      return prefix == ''
        || typeof prefix == "object"
        || r.Prefix.toLowerCase() == prefix.toLowerCase()
    })
  }

  // delete attachment from array
  delete(n:any[string]) {
    if (n.ServerRelativeUrl != null) {
      n.Deleted = true;
    } else {
      this.form[this.field].results = this.form[this.field].results.filter((f:any[string]) => {
        return f.FileName.toLowerCase() != n.FileName.toLowerCase()
      })
    }
  }

  undelete(a:any[string]) {
    a.Deleted = null;
  }

  // add attachment to array
  add(file:any) {
    var prefix = ((this.get('Prefix') || '') != '' && typeof this.get('Prefix') == "string" ? this.get('Prefix') + '-' : '');
    var r = new RegExp(`^${prefix}`, 'i');
    var files:any[any[any]] = [], dup:any[string] = [];

    // ensure the file name doenst already exist as duplicaes are not allowed
    for (var f = 0; f < file.files.length; f++) {
      var n = prefix + file.files[f].name.replace(r, '').replace(/[%'#]/g,'-');
      
      this.form[this.field].results.forEach((a:any[string]) => {
        if (!a.Deleted && a.FileName.toLowerCase() == n.toLowerCase()) {
          dup.push(n);
        }
      });

      if (dup.indexOf(n) == -1)
        files.push(file.files[f]);
    }

    if (dup.length > 0)
      alert('File(s) already exist with name(s): ' + dup.join(', '));
    
    var ths = this;
    files.forEach((f:any) => {
      var reader = new window.FileReader();
      reader.onload = function (event:any) {
        var data = '';
        var bytes = new window.Uint8Array(event.target.result);
        var len = bytes.byteLength;
        for (var i = 0; i < len; i++) {
          data += String.fromCharCode(bytes[i]);
        }

        var n = prefix + f.name.replace(r, '').replace(/[%'#]/g,'-');
        
        ths.form[ths.field].results.push({
          FileName: n,
          ServerRelativeUrl: null,
          Data: data,
          Length: len,
          Prefix: prefix,
          UploadName: n
        });

        setTimeout(() => file.value = null, 10);
      }
      reader.onerror = function () {
        alert('File read error: ' + f.name);
        Logger.log({
          message: `Inside - add(${f.name})`,
          level: LogLevel.Error,
        });
      };
      reader.readAsArrayBuffer(f);
    })
  }

  // prefix passed is array use as drop down
  prefixes(): boolean {
    return typeof this.get('Prefix') == "object"
  }

  // on change prefix drop down
  prefix(a:any, p:any) {
    // clean up file name of all previous prefixes
    var rem = a.FileName;
    this.get('Prefix').forEach((x:any) => {
      var n = ((x || '') != '' ? x + '-' : '');
      var r = new RegExp(`^${n}`, 'i');
      rem = rem.replace(r, '');
    });

    // sufix new prefix value
    rem = (p.value ? p.value + '-' : '') + rem;
    
    // check for duplicates after rename
    var dup:any[string] = [];
    this.form[this.field].results.forEach((a:any[string]) => {
      if (!a.Deleted && a.FileName.toLowerCase() == rem.toLowerCase()) {
        dup.push(rem);
      }
    });

    if (dup.length == 0) {
      // if no duplicates set new file name
      a.FileName = rem;
      a.Prefix = p.value;
    } else {
      // return to last prefix
      p.value = a.Prefix;
      // inform user
      alert('File(s) already exist with name(s): ' + dup.join(', '));
    }
  }

  over(evt:any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.filesOver = true;
  }

  leave(evt:any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.filesOver = false;
  }
  
  drop(evt:any) {
    evt.preventDefault();
    evt.stopPropagation();
    this.add(evt.dataTransfer.files);
    this.filesOver = false;
  }

  // gets the required field properties and/or any overrides
  //declare overrode: any[string];
  get(t:string) {
    var p = null;
    var overrode = this.override ? (typeof this.override == "string" ? JSON.parse(this.override) : this.override) : {};
    if (overrode && overrode[t])
      p = overrode[t];
    if (!p && this.spec && this.spec[this.field.replace(/^OData_/, '')] && this.spec[this.field.replace(/^OData_/, '')][t])
      p = this.spec[this.field.replace(/^OData_/, '')][t];
    if (!p && this.spec && this.spec[this.field] && this.spec[this.field][t])
      p = this.spec[this.field][t];
    // if its a multi choice, ensure the object is the correct type
    if (t == 'TypeAsString' && p == 'MultiChoice' && (!this.form[this.field] || !this.form[this.field].results))
      this.form[this.field] = {
        __metadata: {type: "Collection(Edm.String)"},
        results: []
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
    if (t == 'Choices' && !p)
      return [];
    // if no title use internal field name
    if (t == 'Title' && !p)
      return this.field;
    // if no type, non attached local dev, then use random type
    if (t == 'TypeAsString' && !p) {
      this.spec[this.field] = this.spec[this.field] || {};
      this.spec[this.field][t] = ['Boolean','Choice','Integer','DateTime','Text','Note'].splice(Math.floor(Math.random()*5),1)
      return this.spec[this.field][t];
    }
    return !p || typeof p.results == "undefined" ? p : p.results;
  }

  // text area append only changes needs 1 way bind
  change(event:any) : void {
    this.form[this.field] = event.editor ? event.editor.getContent() : event.target.value;
  }

  // choices need filtering
  choices(): any[string] {
    if (typeof this.filter != "function")
      return this.get('Choices');
    var ths = this;
    return this.get('Choices').filter((x:string) => {
      // on choices exclude the other value
      if (ths.other && ths.other == x)
        return false;
      // if there is a filter use it
      return ths.filter(x, ths);
    });
  }

  // selected field option not in available choices, i.e. other
  notInChoices(): boolean {
    if (!this.form[this.field] || (this.form[this.field] == '-' && this.none))
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

  // on multi select, not using ctrl key
  selChangeC(e:any): void {
    e.stopPropagation();
    let scrollTop = 0;
    if ( e.target.parentNode )
      scrollTop = e.target.parentNode.scrollTop;

    var v = e.target.value.split( '\'' )[1];
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
    setTimeout(( function() { e.target.parentNode.scrollTop = scrollTop; } ), 0 );
    setTimeout(( function() { e.target.parentNode.focus(); } ), 0 );
  }

  // show numbers with only1 dot and without any trailing zeros
  niceNumber(): string {
    return !this.form[this.field] ? '' : this.form[this.field].toString().split('.')[0] + (this.form[this.field].toString().split('.').length == 1 ? '' : '.' + this.form[this.field].toString().split('.')[1].replace(/0*$/,''));
  }

  // required is not needed for hidden/disabled items
  required(): boolean {
    if (!this.get('Required') || this.disabled || this.elRef.nativeElement.hidden)
      return false;
    return true;
  }
}
