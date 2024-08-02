import { spfi, SPFI, SPBrowser } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import * as MSAL from "@azure/msal-browser";
import { Logger, LogLevel } from "@pnp/logging";
import { PermissionKind } from "@pnp/sp/security";
import { PnPLogging } from './PnPLogging';
import { App } from './App';

///<summary>
/// This is to be used in place of specific pnp.sp function when using these form fields to aid in data transforms and a few other fringe cases outlined in the method coments 
///</summary>
export class SharepointChoiceUtils {
    // context can be read and updated
    public context?:string = '';
    public sp:SPFI;

    // attempt to establish correct context url for the site from one of the available sources then setup logging for this class
    constructor(
        context?: string
    ) {
      this.context = context;
      let w:any = window;
      if ((this.context || "") == "")
        this.context = w._spPageContextInfo ? w._spPageContextInfo.webAbsoluteUrl : undefined;
      if ((this.context || "") == "")
        this.context = document.location.href.split('?')[0].split('#')[0].split('/_layouts/')[0].split('/Lists/')[0].split('/Pages/')[0].split('/SitePages/')[0];

      this.context = this.context?.replace(/\/$/,'');

      this.sp = spfi().using(SPBrowser({ baseUrl: this.context }));

      Logger.subscribe(new PnPLogging());
      Logger.activeLogLevel = LogLevel.Warning;

      this.mockClassicContext();
    }

    private async mockClassicContext() {
      let w:any = window;
      // no classic sp context then mock one up
      if (typeof w._spPageContextInfo == "undefined")
        w._spPageContextInfo = {};
      // no user in context or a different web then mock it in
      if (typeof w._spPageContextInfo.userLoginName == "undefined" || w._spPageContextInfo.webAbsoluteUrl != this.context) {
        var user = await this.sp.web.currentUser();
        w._spPageContextInfo = {
          userLoginName: user.LoginName,
          userDisplayName: user.Title,
          userEmail: user.Email,
          userId: user.Id,
          webAbsoluteUrl: this.context,
        };
      }
    }

    // get the current user and permissions to a flat object for easier use in [disabled]="permission['']" etc
    // NOTE: this will only detect direct assignments or users added to a mail enabled global security group
    public async permissions():Promise<any> {
        var p:any = {};
        let w:any = window;

        try {
          await this.mockClassicContext();
          var web = await this.sp.web();
          var webTitle = web.Title;
          var perm = await this.sp.web.currentUser.groups();
          perm.forEach(x => {
            p[x.LoginName] = true;
            if (x.LoginName.startsWith(`${webTitle} `))
              p[x.LoginName.replace(`${webTitle} `,'')] = true;
          })
        } catch (e) {
          p = {Error: true};
        }

        return {userId: w._spPageContextInfo.userId, perms: p}
    }

    // check permission against object
    public async hasPermission(object:any, permissions:any[PermissionKind]):Promise<boolean> {
      try {
        var perm = await object.getCurrentUserEffectivePermissions();
        for (var p in permissions) {
          if (this.sp.web.hasPermissions(perm, permissions[p]))
            return true;
        }
      } catch (e) {}
      return false;
    }
    
    // get list fields in the appropriate format for use in <sharepoint-choice spec=""> attributes
    public async fields(listTitle:string):Promise<any> {
        var spec:any = {'odata.context': this.sp};

        try {
          var arr = await this.sp.web.lists.getByTitle(listTitle).fields();
          arr.forEach(x => {
            spec[x.InternalName] = x;
            // used for people searches only as pnp doesnt have a suitable endpoint yet
            spec[x.InternalName].Context = this.context;
          });
        } catch (e) {
          spec['Title'] = {TypeAsString:'Text',MaxLength:16,Description:'Tooltip'};
        }

        return spec;
    }

    // load list item data and parse any data types appropriate for use in <sharepoint-choice ngModel=""> attributes
    public async data(id:number, listTitle:string):Promise<any> {
        var d:any = {};

        try {
          d = await this.sp.web.lists.getByTitle(listTitle).items.getById(id)();
          for (var key in d) {
            // people fields return twice
            if (key.endsWith('StringId') && (d[key.replace(/StringId$/,'Id')] || d[key.replace(/StringId$/,'Id')] === null))
              delete d[key];

            // if there are attachments start loading
            if (key == 'Attachments') {
              if (d[key] === true)
                d[key] = { results: await this.sp.web.lists.getByTitle(listTitle).items.getById(id).attachmentFiles() };
              else
                d[key] = { results:[] };
            }

            // remove odata. prefixed
            if (key.startsWith('odata.') || key == '__metadata')
              delete d[key];

            // dont process nulls
            if (!d[key] || d[key] === null)
              continue;
              
            // return multifields to results, old behaviour for old people fields and to prevent json paring clashing
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
  
    // calls an api generically
    public async callApi(tenancyOnMicrosoft: string, clientId: string, permissionScope: string, apiUrl?: string, httpMethod?: string, jsonPostData?: any):Promise<any> {
      // client settings
      var config = {
        auth: {
            clientId: clientId,
            authority: `https://login.microsoftonline.com/${tenancyOnMicrosoft}.onmicrosoft.com`,
            redirectUri: this.context?.replace(/\/$/,'')
        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: false
        }
      }
      
      // init client
      var msal = new MSAL.PublicClientApplication(config);

      await msal.initialize();
      
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
        try {
          await msal.loginPopup(params);
          params.account = msal.getAllAccounts()[0];
          login = await msal.acquireTokenSilent(params);
        } catch (e) {
          throw `Exception logging in to MSAL for scope ${permissionScope} with error ${e}`;
        }
      }

      // if no url, login only, then return
      if (apiUrl == null)
        return null;

      // query api
      var r;
      try {
        r = await fetch(apiUrl, {
              method: httpMethod || 'GET',
              headers: {
                  'Authorization': `Bearer ${login.accessToken}`,
                  'Content-Type': 'application/json'
              },
              body: jsonPostData ? JSON.stringify(jsonPostData) : null,
          });
      
        // return formatted data for 2xx, 4xx and 5xx will not return
        if (r.status == 204)
          return null;
        return await r.clone().json();
      } catch (e) {
        throw `Exception getting API data with status ${r?.status} response ${await r?.text()}`;
      }
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
            if ((save[key] === null && uned[key] !== null) || key == "Id")
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
            if (typeof save[key] == "object" && !save[key].results && !save[key].Url)
              save[key] = JSON.stringify(save[key]);
            
            // convert back to direct array and ensure no nulls selected, should never occur but does on some browsers?
            if (typeof save[key] == "object" && save[key].results)
              save[key] = save[key].results.filter((i:any) => i !== null && i !== undefined);
          }
          
          // save/update the item
          if (typeof save.Id == "undefined" || save.Id < 1) {
            var saving = await this.sp.web.lists.getByTitle(listTitle).items.add(save);
            save.Id = saving.Id;
          } else {
            await this.sp.web.lists.getByTitle(listTitle).items.getById(save.Id).update(save);
          }
    
          // process attachments as deletes then uploads
          if (form.Attachments && form.Attachments.results && form.Attachments.results.length > 0) {
            var deletes = form.Attachments.results.filter((a:any) => {
              return a.Deleted
            }).map((a:any) => {
              return a.FileName;
            });

            for (var i = 0; i < deletes.length; i++)
              await this.sp.web.lists.getByTitle(listTitle).items.getById(save.Id).attachmentFiles.getByName(deletes[i]).delete();
    
            var adds = form.Attachments.results.filter((a:any) => {
              return !a.Deleted && !a.ServerRelativeUrl
            }).map((a:any) => {
              return {
                name: a.FileName,
                content: a.Data
              };
            });

            for (var a = 0; a < adds.length; a++) {
              await this.sp.web.lists.getByTitle(listTitle).items.getById(save.Id).attachmentFiles.add(adds[a].name, adds[a].content);
            }
          }
        } catch (e) {
          alert('Error saving');
          throw e;
        }

        return save.Id;
    }

    // get query parameters, not strictly sharepoint but reused a lot
    public param(parameterToReturn:string):string|undefined {
        var rx = new RegExp(`[?&]${parameterToReturn}=([^&]+).*$`);
        var returnVal = document.location.search.match(rx);
        return returnVal === null ? undefined : decodeURIComponent(returnVal[1]).replace(/\+/g, ' ');
    }

    public async ensurePath(path: string, start: number): Promise<void> {
      if (path.indexOf("://") >= 0)
        path = path.substring(path.indexOf('/', 9));

      var p = path.split('/').slice(0, start + 1).join('/');
      var folder = this.sp.web.getFolderByServerRelativePath(p);
      try {
        var f = await folder();
        if (!f.Exists)
          await this.sp.web.getFolderByServerRelativePath(path.split('/').slice(0,start).join('/')).addSubFolderUsingPath(path.split('/').slice(start)[0]);
      } catch (e) {
        await this.sp.web.getFolderByServerRelativePath(path.split('/').slice(0,start).join('/')).addSubFolderUsingPath(path.split('/').slice(start)[0]);
      }
      if (p != path)
        await this.ensurePath(path, start + 1);
    }

    public async getRoot(list:string): Promise<string> {
      let root = await this.sp.web.lists.getByTitle(list).rootFolder();
      return root.ServerRelativeUrl;
    }

    public async getFiles(serverRelative:string, additional:string|undefined): Promise<any> {
      if (serverRelative.indexOf("://") >= 0)
        serverRelative = serverRelative.substring(serverRelative.indexOf('/', 9));

      var files = await this.sp.web.getFolderByServerRelativePath(serverRelative.replace(/\/$/, '') + (additional ? '/'+additional : '')).files.expand('ListItemAllFields')();
      
      var ret = [];
      files.forEach(file => {
        ret.push({
          Name: file.Name,
          FileName: file.Name,
          TimeCreated: file.TimeCreated,
          Classification: file['ListItemAllFields'].Classification,
          OldClassification: file['ListItemAllFields'].Classification,
          Request: file['ListItemAllFields'].Request,
          ServerRelativeUrl: file.ServerRelativeUrl,
          // everything else
          ListItemAllFields: file['ListItemAllFields']
        })
      });

      return ret;
    }
    
    public async saveFiles(path:string, additional:string|undefined, url:any|undefined, files:any, metadata:any|undefined): Promise<void> {
      if (path.indexOf("://") >= 0)
        path = path.substring(path.indexOf('/', 9));

      // common metadata for folder and each file, unless overridden at a file level
      var commonmeta = metadata ? JSON.parse(JSON.stringify(metadata)) : url ? {Request: url} : {};

      var folder = await this.sp.web.getFolderByServerRelativePath(path).getItem();
      await folder.update(commonmeta);
  
      // subfolders for these
      if (additional && additional != '') {
        path += '/' + additional;
        folder = await this.sp.web.getFolderByServerRelativePath(path).getItem();
        await folder.update(commonmeta);
      }
  
      // process saves and deletes
      for (var i = 0; i < files.results.length; i++) {
        var file = files.results[i];
        // clone common metadata for file overrides
        var meta = JSON.parse(JSON.stringify(commonmeta));
        if (file.metadata) {
          for (var m in file.metadata)
            meta[m] = file.metadata[m];
        } else if (file.Classification) {
          meta['Classification'] = file.Classification;
        }
        if (file.Delete)
          await this.sp.web.getFolderByServerRelativePath(path+'/'+file.Name).recycle();
        else if (file.Data) {
          await this.sp.web.getFolderByServerRelativePath(path).files.addUsingPath(file.FileName, file.Data, {Overwrite: true});
          let i = await this.sp.web.getFolderByServerRelativePath(path).files.getByUrl(file.FileName).getItem();
          await i.update(meta);
          // mock the data back in so submit again doesnt fail
          file.Name = file.FileName;
          file.TimeCreated = new Date();
          file.ServerRelativeUrl = path+'/'+file.FileName;
          file.Request = url;
          file.OldClassification = file.Classification;
          file.ListItemAllFields = meta;
          delete file.Data;
        } else if (file.Classification != file.OldClassification || !file.Request) {
          let i = await this.sp.web.getFolderByServerRelativePath(path+'/'+file.Name).getItem();
          await i.update(meta);
        }
      }
    }
}
