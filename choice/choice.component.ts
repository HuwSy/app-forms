import { Component, OnInit, Input, ElementRef } from '@angular/core';

@Component({
  selector: 'app-choice',
  templateUrl: './choice.component.html',
  styleUrls: ['./choice.component.scss']
})
export class ChoiceComponent implements OnInit {
  @Input() form!: any[string]; // form containing field
  @Input() spec!: any[any[string]]; // spec of field loaded from list
  @Input() field!: string; // internal field name on form object, used for push back and against spec
  @Input() override!: string; // override any spec above. sent as string as passing object kills large form performance

  @Input() none!: string; // none option text instead of null
  @Input() other!: string; // Other fill-in option text, will override to allow other
  @Input() filter!: object; // filter choices by a function

  @Input() disabled!: boolean;

  constructor(private elRef: ElementRef) { }

  declare tinymceOptions: object;
  declare tinymceROOptions: object;

  // on init
  ngOnInit(): void {
    // rich text field
    this.tinymceOptions = {
        resize: false,
        selector: "textarea",
        height: 200,
        menubar: false,
        plugins: "textcolor lists table link paste",
        toolbar: "forecolor | bold italic underline | bullist numlist outdent indent | table | link",
        statusbar: false,
        debounce: false,
        paste_data_images: true
    };
    
    // readonly/disabled rich text field
    this.tinymceROOptions = {
        selector: "textarea",
        height: 200,
        menubar: false,
        toolbar: false,
        statusbar: false
    };
  }

  // files post filtering
  attachments(): any[string] {
    if (!this.form[this.field] || !this.form[this.field].results)
      return [];
    return this.form[this.field].results.filter((r:any[string]) => {
      return this.get('Prefix') || '' == ''
        || r.FileName.toLowerCase().indexOf(this.get('Prefix').toLowerCase()+'-') == 0
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
    var files:any[any[any]] = [], dup:any[string] = [], ths = this;
    // ensure the file name doenst already exist as duplicaes are not allowed
    for (var f = 0; f < file.target.files.length; f++) {
      var n = (this.get('Prefix') && this.get('Prefix') != '' ? this.get('Prefix').toLowerCase() + '-' : '');
      if (this.get('Prefix') && this.get('Prefix') != '' && file.target.files[f].name.toLowerCase().indexOf(this.get('Prefix').toLowerCase()+'-') == 0)
        n += file.target.files[f].name.toLowerCase().substring(this.get('Prefix').length + 1).replace(/[%'#]/g,'-');
      else
        n += file.target.files[f].name.toLowerCase().replace(/[%'#]/g,'-');
      
      this.form[this.field].results.forEach((a:any[string]) => {
        if (!a.Deleted && a.FileName.toLowerCase() == n) {
          dup.push(n);
        }
      });

      if (dup.indexOf(n) == -1)
        files.push(file.target.files[f]);
    }

    if (dup.length > 0)
      alert('File(s) already exist with name(s): ' + dup.join(', '));
    
    file.target.value = null;

    files.forEach((f:any) => {
      var reader = new window.FileReader();
      reader.onload = function (event:any) {
        var data = '';
        var bytes = new window.Uint8Array(event.target.result);
        var len = bytes.byteLength;
        for (var i = 0; i < len; i++) {
          data += String.fromCharCode(bytes[i]);
        }

        var n = (ths.get('Prefix') && ths.get('Prefix') != '' ? ths.get('Prefix') + '-' : '');
        if (ths.get('Prefix') && ths.get('Prefix') != '' && f.name.toLowerCase().indexOf(ths.get('Prefix').toLowerCase()+'-') == 0)
          n += f.name.substring(ths.get('Prefix').length + 1).replace(/[%'#]/g,'-');
        else
          n += f.name.replace(/[%'#]/g,'-');
        
        ths.form[ths.field].results.push({
          FileName: n,
          ServerRelativeUrl: null,
          Data: data,
          Length: len
        });
      }
      reader.onerror = function () {
        alert("File reading error " + f.name);
      };
      reader.readAsArrayBuffer(f);
    })
  }

  // gets the required field properties and/or any overrides
  //declare overrode: any[string];
  get(t:string) {
    var p = null;
    //if (!this.overrode && this.override)
    var overrode = this.override ? JSON.parse(this.override) : {};
    if (overrode)
      p = overrode[t];
    if (p == null && this.spec && this.spec[this.field.replace(/^OData_/, '')])
      p = this.spec[this.field.replace(/^OData_/, '')][t];
    if (p == null && this.spec && this.spec[this.field])
      p = this.spec[this.field][t];
    if (p == null && t == 'Choice')
      return [];
    // on choices exclude the other value
    if (p && t == 'Choice')
      p.results = p.results.filter((x:string) => {
        return !this.other || this.other != x
      });
    // if there is a filter use it
    if (p && t == 'Choice' && typeof this.filter == "function")
      return p.results.filter(this.filter);
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
    // if no title use internal field name
    if (!p && t == 'Title')
      return this.field;
    return !p || typeof p.results == "undefined" ? p : p.results;
  }

  // selected field option not in available choices, i.e. other
  notInChoices(): boolean {
    if (!this.form[this.field] || (this.form[this.field] == '-' && this.none))
      return false;
    var choices = this.get('Choices');
    if (typeof this.filter == "function")
      choices = choices.filter(this.filter);
    return choices.filter((x:string) => {
        return x == this.form[this.field];
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
    return this.form[this.field].toString().split('.')[0] + (this.form[this.field].toString().split('.').length == 1 ? '' : '.' + this.form[this.field].toString().split('.')[1].replace(/0*$/,''));
  }

  // required is not needed for hidden/disabled items
  required(): boolean {
    if (!this.get('Required') || this.disabled || this.elRef.nativeElement.hidden)
      return false;
    return true;
  }
}
