import pnp from '@pnp/pnpjs';
import { MsalClient } from "@pnp/msaljsclient";
import {App} from './App';

export class SharepointChoiceUtils {
    constructor(
        private _context: string
    ) {
        pnp.sp.setup({sp:{baseUrl:_context}});
    }

    public async permissions():Promise<any> {
        var p = {}, u = 0;

        try {
          u = await (await pnp.sp.web.currentUser.get()).Id;
          var webTitle = await (await pnp.sp.web.get()).Title;
          var perm = await pnp.sp.web.currentUser.groups.get();
          perm.forEach(x => {
            p[x.LoginName] = true;
            if (x.LoginName.startsWith(`${webTitle} `))
              p[x.LoginName.replace(`${webTitle} `,'')] = true;
          })
        } catch (e) {
          p = {Error: true};
        }

        return {userId: u, perms: p}
    }
    
    public async fields(list):Promise<any> {
        var spec = {};

        try {
            var arr = await pnp.sp.web.lists.getByTitle(list).fields.get();
            arr.forEach(x => {
                spec[x.InternalName] = x;
                spec[x.InternalName].Context = this._context;
            });
        } catch (e) {
            spec['Title'] = {TypeAsString:'Text',MaxLength:16,Description:'Tooltip'};
        }

        return spec;
    }

    public param(q:string):string {
        var rx = new RegExp(`[?&]${q}=([^&]+).*$`);
        var returnVal = document.location.search.match(rx);
        return returnVal === null ? null : returnVal[1];
    }

    public async data(id:number, list):Promise<any> {
        var d = {};

        try {
          d = await pnp.sp.web.lists.getByTitle(list).items.getById(id).get();
          for (var key in d) {
            // people fields return twice
            if (key.endsWith('StringId') && (d[key.replace(/StringId$/,'Id')] || d[key.replace(/StringId$/,'Id')] === null))
              delete d[key];
            // if there are attachments start loading
            if (key == 'Attachments') {
              if (d[key] === true)
                d[key] = { results: await pnp.sp.web.lists.getByTitle(list).items.getById(id).attachmentFiles() };
              else
                d[key] = { results:[] };
            }
            // extract metadata for save
            if (key == 'odata.type')
              d['__metadata'] = {type:d[key]};
            // remove odata. prefixed
            if (key.startsWith('odata.'))
              delete d[key];
            // dont process nulls
            if (!d[key])
              continue;
            // parse objects
            try {
              if (d[key].toString().trim().substring(0,1) == '{' || d[key].toString().trim().substring(0,1) == '[') {
                d[key] = JSON.parse(d[key]);
                d[key] = this.parseLoop(d[key]);
                continue;
              }
            } catch (e) {}
            // dates
            if (d[key].toString().match(/[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]/) != null) {
              d[key] = new Date(d[key]);
              continue;
            }
          }
        } catch (e) {
          alert('Error loading');
          throw e;
        }

        return d;
    }

    private parseLoop(i:any):any {
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
  
    public async history(id:number, list):Promise<any> {
        return await pnp.sp.web.lists.getByTitle(list).items.getById(id).versions.get();
    }

    public async msalApi(clientId: string, tokenRole: string, endPoint: string, release: string):Promise<any> {
        var c = new MsalClient({
          auth: {
              authority: `https://login.microsoftonline.com/${App.Tenancy}.onmicrosoft.com`,
              clientId: clientId,
              redirectUri: this._context
          }
        });
        var t = await c.getToken([`${App.Token}/${release}/${tokenRole}`]);
        
        var r = await fetch(~document.location.href.toLowerCase().indexOf('workbench.aspx') || document.location.host.toLowerCase().startsWith('localhost')
            ? `https://localhost/${endPoint}`
            : `${App.Token}/${release}/${endPoint}`, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${t}`
                }
            });
        
        return await r.json();
    }

    public async save(form, uned, list):Promise<number> {
        var save = JSON.parse(JSON.stringify(form));

        try {
          delete save["$$hashKey"];

          for (var key in save) {
            if (save[key] === null || key == "Id" || key == "__metadata")
              continue;
            
            // remove and unedited, including internal fields
            if (key == "Attachments" || ((uned[key] || uned[key] === null) && JSON.stringify(uned[key]) == JSON.stringify(save[key]))) {
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
          
          // save/update the item
          if (typeof form.Id == "undefined" || form.Id < 1) {
            save.Id = await (await pnp.sp.web.lists.getByTitle(list).items.add(save)).data.Id;
          } else {
            await pnp.sp.web.lists.getByTitle(list).items.getById(form.Id).update(save);
          }
    
          // process attachments as deletes then uploads
          if (form.Attachments && form.Attachments.results && form.Attachments.results.length > 0) {
            var deletes = form.Attachments.results.filter(a => {
              return a.Deleted
            }).map(a => {
              return a.FileName;
            });

            //if (deletes.length > 0)
            //  await pnp.sp.web.lists.getByTitle(this.list).items.getById(this.form.Id).attachmentFiles.deleteMultiple.apply(deletes);
            for (var i = 0; i < deletes.length; i++)
              await pnp.sp.web.lists.getByTitle(list).items.getById(form.Id).attachmentFiles.getByName(deletes[i]).delete();
    
            var adds = form.Attachments.results.filter(a => {
              return !a.Deleted && !a.ServerRelativeUrl
            }).map(a => {
              return {
                name: a.FileName,
                content: a.Data
              };
            });

            if (adds.length > 0)
              await pnp.sp.web.lists.getByTitle(list).items.getById(form.Id).attachmentFiles.addMultiple(adds);
          }
        } catch (e) {
          alert('Error saving');
          throw e;
        }

        return save.Id;
    }
}
