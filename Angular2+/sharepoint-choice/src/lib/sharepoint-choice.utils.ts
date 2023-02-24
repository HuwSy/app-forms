import pnp from '@pnp/pnpjs';
import * as MSAL from "@azure/msal-browser";
import { Logger, LogLevel } from "@pnp/logging";
import { PnPLogging } from './PnPLogging';
import { App } from './App';

///<summary>
/// This is to be used in place of specific pnp.sp function when using these form fields to aid in data transforms and a few other fringe cases outlined in the method coments 
///</summary>
export class SharepointChoiceUtils {
    // context can be read and updated
    public context:string = '';

    // attempt to establish correct context url for the site from one of the available sources then setup logging for this class
    constructor(
        context?: string
    ) {
      this.context = context;
      if ((this.context || "") == "")
        this.context = typeof window['_spPageContextInfo'] == "object" ? window['_spPageContextInfo']['webAbsoluteUrl'] : null;
      if ((this.context || "") == "")
        this.context = document.location.href.split('?')[0].split('#')[0].split('/_layouts/')[0].split('/Lists/')[0].split('/Pages/')[0].split('/SitePages/')[0];

      this.context = this.context.replace(/\/$/,'');

      pnp.sp.setup({sp:{baseUrl:this.context}});
      Logger.subscribe(new PnPLogging());
      Logger.activeLogLevel = LogLevel.Warning;
    }

    // get the current user and permissions to a flat object for easier use in [disabled]="permission['']" etc
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
    
    // get list fields in the appropriate format for use in <sharepoint-choice spec=""> attributes
    public async fields(listTitle:string):Promise<any> {
        var spec = {'odata.metadata': this.context};

        try {
            var arr = await pnp.sp.web.lists.getByTitle(listTitle).fields.get();
            arr.forEach(x => {
                spec[x.InternalName] = x;
                spec[x.InternalName].Context = this.context;
            });
        } catch (e) {
            spec['Title'] = {TypeAsString:'Text',MaxLength:16,Description:'Tooltip'};
        }

        return spec;
    }

    // load list item data and parse any data types appropriate for use in <sharepoint-choice ngModel=""> attributes
    public async data(id:number, listTitle:string):Promise<any> {
        var d = {};

        try {
          d = await pnp.sp.web.lists.getByTitle(listTitle).items.getById(id).get();
          for (var key in d) {
            // people fields return twice
            if (key.endsWith('StringId') && (d[key.replace(/StringId$/,'Id')] || d[key.replace(/StringId$/,'Id')] === null))
              delete d[key];

            // if there are attachments start loading
            if (key == 'Attachments') {
              if (d[key] === true)
                d[key] = { results: await pnp.sp.web.lists.getByTitle(listTitle).items.getById(id).attachmentFiles() };
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
            if (!d[key] || d[key] === null)
              continue;
              
            // return arrays back to results, fix pnpjs not behaving as expected
            if (typeof d[key] == "object" && !d[key].results && d[key].length > 0) {
              d[key] = {
                results: d[key],
                __metadata: {type: (typeof d[key][0] == "number" ? "Collection(Edm.Int32)" : "Collection(Edm.String)")}
              }
            }

            // parse objects within text fields for looped data
            try {
              let f = d[key].toString().trim().substring(0,1);
              if ((f == '{' || f == '[') && d[key].toString().trim().endsWith(f == '{' ? '}' : ']')) {
                d[key] = JSON.parse(d[key]);
                d[key] = this.parseLoop(d[key]);
                continue;
              }
            } catch (e) {}

            // dates and date times
            let i = d[key].toString();
            if (i.match(/^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9][ T][0-2][0-9]:[0-5][0-9]:*[0-9]*\.*[0-9]*Z*$/) != null
              || i.match(/^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]$/) != null) {
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
          if (i.match(/^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9][ T][0-2][0-9]:[0-5][0-9]:*[0-9]*\.*[0-9]*Z*$/) != null
            || i.match(/^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]$/) != null) {
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
  
    // calls an api more generically
    public async callApi(tenancyOnMicrosoft: string, clientId: string, permissionScope: string, apiUrl: string, httpMethod: string, jsonPostData: any):Promise<any> {
      // client settings
      var config = {
        auth: {
            clientId: clientId,
            authority: `https://login.microsoftonline.com/${tenancyOnMicrosoft}.onmicrosoft.com`,
            redirectUri: this.context.replace(/\/$/,'')
        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: false
        }
      }
      
      // init client
      var msal = new MSAL.PublicClientApplication(config);
      
      // permission settings
      var params = {
        scopes: [permissionScope],
        account: msal.getAllAccounts()[0]
      };

      // attempt to get token or login and get token
      var login;
      try {
        login = await msal.acquireTokenSilent(params);
      } catch (error) {
        await msal.loginPopup(params);
        params.account = msal.getAllAccounts()[0];
        login = await msal.acquireTokenSilent(params);
      }

      // query api
      var r = await fetch(apiUrl, {
              method: httpMethod,
              headers: {
                  'Authorization': `Bearer ${login.accessToken}`,
                  'Content-Type': 'application/json'
              },
              body: jsonPostData ? JSON.stringify(jsonPostData) : null,
          });
      
      // return formatted data
      return await r.json();
    }

    // patch save list item data and parse any data types appropriate for use in <sharepoint-choice ngModel=""> attributes
    public async save(formDataIncIdToUpdate: any, uneditedDataToBuildPatch: any, listTitle: string):Promise<number> {
        let form = formDataIncIdToUpdate, uned = uneditedDataToBuildPatch;
        var save = JSON.parse(JSON.stringify(form));
        if (!uned || uned == null)
          uned = {};

        try {
          delete save["$$hashKey"];

          for (var key in save) {
            if ((save[key] === null && uned[key] !== null) || key == "Id" || key == "__metadata")
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
            var saving = await pnp.sp.web.lists.getByTitle(listTitle).items.add(save);
            save.Id = saving.data.Id;
          } else {
            await pnp.sp.web.lists.getByTitle(listTitle).items.getById(save.Id).update(save);
          }
    
          // process attachments as deletes then uploads
          if (form.Attachments && form.Attachments.results && form.Attachments.results.length > 0) {
            var deletes = form.Attachments.results.filter(a => {
              return a.Deleted
            }).map(a => {
              return a.FileName;
            });

            for (var i = 0; i < deletes.length; i++)
              await pnp.sp.web.lists.getByTitle(listTitle).items.getById(save.Id).attachmentFiles.getByName(deletes[i]).delete();
    
            var adds = form.Attachments.results.filter(a => {
              return !a.Deleted && !a.ServerRelativeUrl
            }).map(a => {
              return {
                name: a.FileName,
                content: a.Data
              };
            });

            if (adds.length > 0)
              await pnp.sp.web.lists.getByTitle(listTitle).items.getById(save.Id).attachmentFiles.addMultiple(adds);
          }
        } catch (e) {
          alert('Error saving');
          throw e;
        }

        return save.Id;
    }

    // get query parameters, not strictly sharepoint but reused a lot
    public param(parameterToReturn:string):string {
        var rx = new RegExp(`[?&]${parameterToReturn}=([^&]+).*$`);
        var returnVal = document.location.search.match(rx);
        return returnVal === null ? null : returnVal[1];
    }
}
