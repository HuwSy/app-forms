import { Component, OnInit, ViewEncapsulation } from '@angular/core';
import pnp from '@pnp/pnpjs';

@Component({
  selector: 'app-sp-fx-aweb-part',
  templateUrl: './sp-fx-aweb-part.component.html',
  styleUrls: ['./sp-fx-aweb-part.component.scss'],
  encapsulation: ViewEncapsulation.Emulated
})
export class SpFxAWebPartComponent implements OnInit {
  declare form:any[string];
  declare uned:any[string];
  declare spec:any[string];
  declare perm:any[string];

  constructor() { }

  ngOnInit(): void {
    this.spec = {};
    this.form = {__metadata:{type:'SP.Data.NullListItem'}};

    this.perm = {Test: true};

    this.fields();
    this.data(1);
  }

  async fields():Promise<void> {
    try {
      var arr = await pnp.sp.web.lists.getByTitle('null').fields.get();
      arr.forEach(x => {
        this.spec[x.InternalName] = x;
      });
    } catch (e) {
      this.spec = {URL:{TypeAsString:'URL',Title:'Url',Required:true},Title:{TypeAsString:'Text',MaxLength:16,Description:'Tooltip'},People:{TypeAsString:'UserMulti'}};
    }
  }

  parseLoop(i:any) {
    try {
      if (typeof i == "string") {
        if (i.match(/20[0-9]{2}\-[01][0-9]\-[0-3][0-9]/) != null) {
          return new Date(i);
        }
      } else if (typeof i == "object") {
        try {
          for (var a in i)
            i[a] = this.parseLoop(i[a]);
        } catch (e) {}
      } 
    } catch (e) {}
    return i;
  }

  async data(id:number):Promise<void> {
    try {
      this.form = await pnp.sp.web.lists.getByTitle('null').items.getById(id).get();
      for (var key in this.form) {
        // people fields return twice
        if (key.endsWith('StringId') && (this.form[key.replace(/StringId$/,'Id')] || this.form[key.replace(/StringId$/,'Id')] === null)) {
          delete this.form[key];
          continue;
        }
        // parse objects
        try {
					if (this.form[key].toString().trim().substring(0,1) == '{' || this.form[key].toString().trim().substring(0,1) == '[') {
						this.form[key] = JSON.parse(this.form[key]);
						this.form[key] = this.parseLoop(this.form[key]);
						continue;
					}
				} catch (e) {}
				// dates
				if (this.form[key].toString().match(/[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]/) != null) {
					this.form[key] = new Date(this.form[key]);
					continue;
				}
      }
      this.uned = JSON.parse(JSON.stringify(this.form));
    } catch (e) {
      this.form.URL = {Url: 'https://d', Description: 'D'};
    }
  }

  async save():Promise<void> {
    try {
      var save = JSON.parse(JSON.stringify(this.form));
      delete save["$$hashKey"];
      for (var key in save) {
        if (save[key] === null || key == "Id" || key == "__metadata")
          continue;
        // remove and unedited, including internal fields
        if ((this.uned[key] || this.uned[key] === null) && JSON.stringify(this.uned[key]) == JSON.stringify(save[key])) {
          delete save[key];
          continue;
        }
        // convert dates
        if (typeof save[key].toJSON != "undefined") {
          save[key] = save[key].toJSON();
          continue;
        }
        // convert JSON
        if (typeof save[key] == "object" && !save[key].results)
          save[key] = JSON.stringify(save[key]);
      }
      
      if (typeof this.form.Id == "undefined" || this.form.Id < 1) {
        pnp.sp.web.lists.getByTitle('null').items.add(save);
      } else {
        pnp.sp.web.lists.getByTitle('null').items.getById(this.form.Id).update(save);
      }
    } catch (e) {
      alert('Error saving');
    }
  }
}
