import pnp from '@pnp/pnpjs';
import * as MSAL from "@azure/msal-browser";
import {App} from './App';

export class SharepointChoiceUtils {
    public context:string = '';

    constructor(
        c: string
    ) {
        if ((c || "") == "") {
          if (typeof window['_spPageContextInfo'] == "object") {
            this.context = window['_spPageContextInfo']['webAbsoluteUrl'];
          } else {
            this.context = document.location.href.split('?')[0].split('#')[0].split('/_layouts/')[0].split('/Lists/')[0].split('/Pages/')[0].split('/SitePages/')[0];
          }
        } else {
          this.context = c;
        }

        pnp.sp.setup({sp:{baseUrl:this.context}});
    }

    public async permissions():Promise<any> {
        var p = {}, u = 0;

        try {
          var user = await pnp.sp.web.currentUser.get();
          u = user.Id;
          var web = await pnp.sp.web.get();
          var webTitle = web.Title;
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
                spec[x.InternalName].Context = this.context;
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
      // connect client
      var config = {
        auth: {
            clientId: clientId,
            authority: `https://login.microsoftonline.com/${App.Tenancy}.onmicrosoft.com`,
            redirectUri: this.context.replace(/\/$/,'')
        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: false
        }
      }
      
      var msal = new MSAL.PublicClientApplication(config);
      
      var params = {
        scopes: [`${App.Token}/${release}/${tokenRole}`],
        account: msal.getAllAccounts()[0]
      };

      var login;
      try {
        login = await msal.acquireTokenSilent(params);
      } catch (error) {
        await msal.loginPopup(params);
        params.account = msal.getAllAccounts()[0];
        login = await msal.acquireTokenSilent(params);
      }

      // query api
      var r = await fetch(release == 'Local'
          ? `https://localhost/${endPoint}`
          : `${App.Token}/${release}/${endPoint}`, {
              method: 'GET',
              headers: {
                  'Authorization': `Bearer ${login.accessToken}`
              }
          });
      
      // return formatted data
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
          if (typeof save.Id == "undefined" || save.Id < 1) {
            var saving = await pnp.sp.web.lists.getByTitle(list).items.add(save);
            save.Id = saving.data.Id;
          } else {
            await pnp.sp.web.lists.getByTitle(list).items.getById(save.Id).update(save);
          }
    
          // process attachments as deletes then uploads
          if (form.Attachments && form.Attachments.results && form.Attachments.results.length > 0) {
            var deletes = form.Attachments.results.filter(a => {
              return a.Deleted
            }).map(a => {
              return a.FileName;
            });

            for (var i = 0; i < deletes.length; i++)
              await pnp.sp.web.lists.getByTitle(list).items.getById(save.Id).attachmentFiles.getByName(deletes[i]).delete();
    
            var adds = form.Attachments.results.filter(a => {
              return !a.Deleted && !a.ServerRelativeUrl
            }).map(a => {
              return {
                name: a.FileName,
                content: a.Data
              };
            });

            if (adds.length > 0)
              await pnp.sp.web.lists.getByTitle(list).items.getById(save.Id).attachmentFiles.addMultiple(adds);
          }
        } catch (e) {
          alert('Error saving');
          throw e;
        }

        return save.Id;
    }
}
