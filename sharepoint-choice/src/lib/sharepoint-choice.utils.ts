import { spfi, SPFI, SPBrowser } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import "@pnp/sp/site-groups";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { PublicClientApplication } from "@azure/msal-browser";
import { PermissionKind } from "@pnp/sp/security";
import { App } from './App';

///<summary>
/// This is to be used in place of specific pnp.sp function when using these form fields to aid in data transforms and a few other fringe cases outlined in the method coments 
///</summary>
export class SharepointChoiceUtils {
  // context can be read and updated
  public context?: string = '';
  public sp: SPFI;

  // attempt to establish correct context url for the site from one of the available sources then setup logging for this class
  constructor(
    context?: string
  ) {
    this.context = context;
    let w: any = window;
    if (!this.context)
      this.context = (w._spPageContextInfo ? w._spPageContextInfo.webAbsoluteUrl : null) ?? document.location.href.split('?')[0].split('#')[0].split('/_layouts/')[0].split('/Lists/')[0].split('/Pages/')[0].split('/SitePages/')[0];

    this.context = this.context?.replace(/\/$/, '');

    this.sp = spfi().using(SPBrowser({ baseUrl: this.context }));

    this.mockClassicContext();
    this.watermark();
  }

  private watermark() {
    if (App.Release != 'LIVE' && document.getElementById('sp-environment-watermark') == null) {
      var environment = document.createElement('style');
      environment.id = 'sp-environment-watermark';
      environment.innerHTML = `[ng-version]::before{content: '${Array(11).fill(App.Release).join(' - ')}';}`;
      document.head.appendChild(environment);
    }
  }

  private async mockClassicContext() {
    let w: any = window;
    // no classic sp context then mock one up
    if (!w._spPageContextInfo)
      w._spPageContextInfo = {};
    // no user in context or a different web then mock it in
    if (!w._spPageContextInfo.userLoginName || w._spPageContextInfo.webAbsoluteUrl != this.context) {
      var web = await this.sp.web();
      var user = await this.sp.web.currentUser();
      w._spPageContextInfo = {
        userLoginName: user.LoginName,
        userDisplayName: user.Title,
        userEmail: user.Email,
        userId: user.Id,
        webAbsoluteUrl: this.context,
        webTitle: web.Title,
      };
    }
  }

  // get the current user and permissions to a flat object for easier use in [disabled]="permission['']" etc
  // NOTE: this will only detect direct assignments or users added to a mail enabled global security group
  public async permissions(): Promise<any> {
    var p: any = {};
    let w: any = window;

    try {
      var web = await this.sp.web();

      // get any directly assigned groups
      // this doesnt work well with ad and aad groups assignments
      var perm = await this.sp.web.currentUser.groups();
      perm.forEach(x => {
        p[x.LoginName] = true;
        if (x.LoginName.startsWith(`${web.Title} `))
          p[x.LoginName.replace(`${web.Title} `, '')] = true;
      });

      // ad and aad groups within sp groups dont always expose groups above
      // this depends on hidden, no crawl list with specific permissions assigned to the same list item title and created by SHAREPOINT\System Account
      try {
        var sec = await this.sp.web.lists.getByTitle('Security')();
        if (sec.Hidden && sec.IsApplicationList) {
          var per = await this.sp.web.lists.getByTitle('Security').items.select("Title").top(5000)();
          per.forEach(s => {
            p[s.Title] = true;
            if (s.Title.startsWith(`${web.Title} `))
              p[s.Title.replace(`${web.Title} `, '')] = true;
          })
        }
      } catch (e) { }
    } catch (e) {
      p = { Error: true };
    }

    return { userId: w._spPageContextInfo.userId, perms: p }
  }

  // check permission against object
  public async hasPermission(object: any, permissions: any[PermissionKind]): Promise<boolean> {
    try {
      var perm = await object.getCurrentUserEffectivePermissions();
      for (var p in permissions) {
        if (this.sp.web.hasPermissions(perm, permissions[p]))
          return true;
      }
    } catch (e) { }
    return false;
  }

  // get list fields in the appropriate format for use in <sharepoint-choice spec=""> attributes
  public async fields(listTitle: string): Promise<any> {
    var spec: any = { 'odata.context': this.sp };

    try {
      // even though the main fields are in the selection not all are returned such as Format, so parse the SchemaXml for the rest
      var selectFields = 'Title,InternalName,TypeAsString,Required,Choices,MaxLength,Description,DisplayFormat,AppendOnly,SelectionGroup,Format,FillInChoice,RichText,ReadOnlyField,DefaultValue,SchemaXml'.split(',');
      var arr = await this.sp.web.lists.getByTitle(listTitle).fields.select(...selectFields)();
      arr.forEach(x => {
        if (x.SchemaXml) {
          let s = (new DOMParser()).parseFromString(x.SchemaXml, "text/xml").documentElement.attributes;
          Array.from(s).reduce((acc, attr) => {
            if (!x[attr.name]) {
              x[attr.name] = attr.value;
              x[attr.name] = x[attr.name] == 'TRUE' ? true : x[attr.name] == 'FALSE' ? false : x[attr.name];
              x[attr.name] = x[attr.name] != null && !isNaN(parseFloat(x[attr.name])) && isFinite(x[attr.name]) ? parseFloat(x[attr.name]) : x[attr.name];
            }
            return acc;
          }, {});
          // prevent reparsing anywhere else
          x.SchemaXml = "";
        }
        spec[x.InternalName] = x;
      });
    } catch (e) {
      spec['Title'] = { TypeAsString: 'Text', MaxLength: 16, Description: 'Tooltip' };
    }

    return spec;
  }

  private async cleanLoadKeys(d: any, listTitle?: string, id?: number) {
    for (var key in d) {
      // people fields return twice
      if (key.endsWith('StringId') && (d[key.replace(/StringId$/, 'Id')] || d[key.replace(/StringId$/, 'Id')] === null)) {
        delete d[key];
        continue;
      }

      // if there are attachments start loading
      if (key == 'Attachments' && listTitle && id) {
        if (d[key] === true)
          d[key] = { results: await this.sp.web.lists.getByTitle(listTitle).items.getById(id).attachmentFiles() };
        else
          d[key] = { results: [] };
        continue;
      }

      // remove odata. prefixed
      if (key.startsWith('odata.') || key == '__metadata') {
        delete d[key];
        continue;
      }

      // blank is null
      if (d[key] === '')
        d[key] = null;

      // dont process nulls or blanks
      if (d[key] == null)
        continue;

      // return multifields to results, old behaviour for old people fields and to prevent json paring clashing
      if (typeof d[key] == "object" && !d[key].results && d[key].length > 0) {
        d[key] = {
          results: d[key],
          __metadata: { type: (typeof d[key][0] == "number" ? "Collection(Edm.Int32)" : "Collection(Edm.String)") }
        }
        continue;
      }

      // parse objects within text fields for looped data
      try {
        let f = d[key].toString().trim().substring(0, 1);
        if ((f == '{' || f == '[') && d[key].toString().trim().endsWith(f == '{' ? '}' : ']')) {
          d[key] = JSON.parse(d[key]);
          d[key] = this.parseLoop(d[key]);
          continue;
        }
      } catch (e) { }

      // dates and date times
      let i = d[key].toString();
      if (i.match(/^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9][ T][0-2][0-9]:[0-5][0-9]:*[0-9]*\.*[0-9]*Z*$/) != null
        || i.match(/^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]$/) != null) {
        d[key] = new Date(d[key]);
        continue;
      }
    }
  }

  // load list item data and parse any data types appropriate for use in <sharepoint-choice ngModel=""> attributes
  public async data(id: number, listTitle: string): Promise<any> {
    var d: any = {};

    try {
      d = await this.sp.web.lists.getByTitle(listTitle).items.getById(id)();
      await this.cleanLoadKeys(d, listTitle, id);
    } catch (e) {
      alert('Error loading:\n\n' + e);
      throw e;
    }

    return d;
  }

  private parseLoop(i: any): any {
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
        } catch (e) { }
      }
    } catch (e) { }
    return i;
  }

  // use mappings to determine the api to call, then call it with the correct parameters, with the ability to override to localhost api
  public async msalApi(serverRelativeEndPoint: string, tokenRole: string, httpMethod: string = 'GET', jsonPostData: any = null, dataType: string = 'json', environment: string = App.Release): Promise<any> {
    // use mappings to determine the api to call
    var endPoint = serverRelativeEndPoint.replace(/^\//, '');
    var api = (App.ApiMap || {})[endPoint.split('/')[0].toLowerCase()];

    return this.callApi(
      App.Tenancy,
      api?.[environment] || api?.['DEV'],
      `${App.ApiToken?.[api?.server]?.[environment] ?? App.ApiToken?.[api?.server]?.['DEV'] ?? ''}${api?.name}/${tokenRole}`,
      endPoint.split('/').length == 1 ? undefined
        : environment == 'LOCAL' || !App.ApiServers?.[api?.server]?.[environment] ? `https://localhost:${api?.port || 44301}/${endPoint}` : `${App.ApiServers[api?.server][environment]}/${endPoint}`,
      httpMethod,
      jsonPostData,
      dataType
    );
  }

  // calls an api more generically, or graph api if no parameters passed
  public async callApi(tenancyOnMicrosoft?: string, clientId?: string, permissionScope?: string, apiUrl?: string, httpMethod?: string, jsonPostData?: any, dataType: string = 'json'): Promise<any> {
    // client settings
    var config = {
      auth: {
        clientId: clientId || App.GraphClient,
        authority: `https://login.microsoftonline.com/${tenancyOnMicrosoft || App.Tenancy}.onmicrosoft.com`,
        redirectUri: this.context?.replace(/\/$/, '')
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
      }
    }

    // init client
    var msal = new PublicClientApplication(config);

    await msal.initialize();

    // permission settings
    var params = {
      scopes: permissionScope ? [permissionScope] : [],
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
    if (!apiUrl)
      return null;

    // query api
    var r;
    try {
      r = await fetch(apiUrl, {
        method: httpMethod || 'GET',
        headers: {
          'Authorization': `Bearer ${login.accessToken}`,
          'Content-Type': dataType == 'json' ? 'application/json' : ''
        },
        body: jsonPostData ? JSON.stringify(jsonPostData) : null,
      });

      // return formatted data for 2xx, 4xx and 5xx will not return
      if (r.status == 204) return null;
      if (r.status < 200 || r.status > 299) throw 'Exception';

      if (dataType == 'json') return await r.json();
      if (dataType == 'text') return await r.text();
      if (dataType == 'buffer') return await r.arrayBuffer();
      if (dataType && dataType != 'none') return new Blob([await r.arrayBuffer()], { type: dataType });

      return r;
    } catch (e) {
      throw `Exception getting API data with status ${r?.status} response ${e} and body ${r?.body}`;
    }
  }

  private cleanSaveKeys(save: any, uned: any): any {
    if (!uned)
      uned = {};

    delete save["$$hashKey"];

    for (var key in save) {
      if (save[key] === '')
        save[key] = null;
      
      if ((save[key] === null && uned[key] !== null) || key == "Id")
        continue;

      // remove and unedited, including internal fields
      if (key == "Attachments" || (uned[key] !== undefined && JSON.stringify(uned[key]) == JSON.stringify(save[key]))) {
        delete save[key];
        continue;
      }

      // prevent errors on nulls
      if (save[key] == null)
        continue;

      // convert dates
      if (save[key].toJSON) {
        save[key] = save[key].toJSON();
        continue;
      }

      // convert JSON
      if (typeof save[key] == "object" && !save[key].results && !save[key].Url) {
        save[key] = JSON.stringify(save[key]);
        continue;
      }

      // convert back to direct array and ensure no nulls selected, should never occur but does on some browsers?
      if (typeof save[key] == "object" && save[key].results) {
        save[key] = save[key].results.filter((i: any) => i);
        continue;
      }
    }
  }

  private hasData(save) : boolean {
    for (var key in save)
      if (key != "Id")
        return true;

    return false;
  }

  // patch save list item data and parse any data types appropriate for use in <sharepoint-choice ngModel=""> attributes
  public async save(formDataIncIdToUpdate: any, uneditedDataToBuildPatch: any, listTitle: string): Promise<number> {
    var save = JSON.parse(JSON.stringify(formDataIncIdToUpdate));
    var errors: Array<string> = [];
    try {
      this.cleanSaveKeys(save, uneditedDataToBuildPatch);

      // save/update the item
      if (!save.Id) {
        var saving = await this.sp.web.lists.getByTitle(listTitle).items.add(save);
        save.Id = saving.Id;
      } else if (this.hasData(save)) {
        await this.sp.web.lists.getByTitle(listTitle).items.getById(save.Id).update(save);
      }

      // process attachments as deletes then uploads
      if (formDataIncIdToUpdate.Attachments && formDataIncIdToUpdate.Attachments.results && formDataIncIdToUpdate.Attachments.results.length > 0) {
        var deletes = formDataIncIdToUpdate.Attachments.results.filter((a: any) => {
          return a.Deleted
        }).map((a: any) => {
          return a.FileName;
        });

        for (var i = 0; i < deletes.length; i++)
          try {
            await this.sp.web.lists.getByTitle(listTitle).items.getById(save.Id).attachmentFiles.getByName(deletes[i]).delete();
          } catch (e) {
            errors.push(`Error deleting attachment ${deletes[i]} for item ${save.Id} in list ${listTitle} with error ${e}`);
          }

        var adds = formDataIncIdToUpdate.Attachments.results.filter((a: any) => {
          return !a.Deleted && !a.ServerRelativeUrl
        }).map((a: any) => {
          return {
            name: a.FileName,
            content: a.Data
          };
        });

        for (var a = 0; a < adds.length; a++)
          try {
            await this.sp.web.lists.getByTitle(listTitle).items.getById(save.Id).attachmentFiles.add(adds[a].name, adds[a].content);
          } catch (e) {
            errors.push(`Error adding attachment ${adds[a].name} for item ${save.Id} in list ${listTitle} with error ${e}`);
          }
      }
    } catch (e) {
      alert('Error saving data:\n\n' + e);
      throw e;
    }

    if (errors.length > 0) {
      alert('Error saving attachments:\n\n' + errors.join('\n'));
      throw errors.join('\n');
    }

    return save.Id;
  }

  // get query parameters, not strictly sharepoint but reused a lot
  public param(parameterToReturn: string): string | undefined {
    var rx = new RegExp(`[?&]${parameterToReturn}=([^&]+).*$`);
    var returnVal = document.location.search.match(rx);
    return !returnVal ? undefined : decodeURIComponent(returnVal[1]).replace(/\+/g, ' ');
  }

  public async ensurePath(path: string, start: number): Promise<void> {
    if (path.indexOf("://") >= 0)
      path = path.substring(path.indexOf('/', 9));
    path = decodeURIComponent(path).replace(/\/$/, '');

    var p = path.split('/').slice(0, start + 1).join('/');
    var folder = this.sp.web.getFolderByServerRelativePath(p);
    try {
      var f = await folder();
      if (!f.Exists)
        await this.sp.web.getFolderByServerRelativePath(path.split('/').slice(0, start).join('/')).addSubFolderUsingPath(path.split('/').slice(start)[0]);
    } catch (e) {
      await this.sp.web.getFolderByServerRelativePath(path.split('/').slice(0, start).join('/')).addSubFolderUsingPath(path.split('/').slice(start)[0]);
    }
    if (p != path)
      await this.ensurePath(path, start + 1);
  }

  public async getRoot(list: string): Promise<string> {
    let root = await this.sp.web.lists.getByTitle(list).rootFolder();
    return root.ServerRelativeUrl;
  }

  public async getFiles(path: string, additional: string | undefined): Promise<any> {
    if (path.indexOf("://") >= 0)
      path = path.substring(path.indexOf('/', 9));
    path = decodeURIComponent(path).replace(/\/$/, '');

    var files = await this.sp.web.getFolderByServerRelativePath(path + (additional ? '/' + additional : '')).files.orderBy('TimeCreated').expand('ListItemAllFields')();

    var ret: Array<any> = [];
    files.forEach(async (file: any) => {
      await this.cleanLoadKeys(file['ListItemAllFields']);

      ret.push({
        FileName: file.Name,
        TimeCreated: file.TimeCreated,
        ServerRelativeUrl: file.ServerRelativeUrl,
        ListItemAllFields: file['ListItemAllFields'],
        OldListItemAllFields: JSON.parse(JSON.stringify(file['ListItemAllFields']))
      })
    });

    return ret;
  }

  public async relocateFolder(source: string, destination: string): Promise<string|null> {
    // ensure these are server relative paths
    var dst = decodeURIComponent(destination.includes("://") ? destination.substring(destination.indexOf("/", 9)) : destination);
    var src = decodeURIComponent(source.includes("://") ? source.substring(source.indexOf("/", 9)) : source);

    // if the destination folder is the same as the current then return null
    if (src.toLowerCase().replace(/\/$/, '') == dst.toLowerCase().replace(/\/$/, ''))
      return null;

    // move the files to the new folder by making the parent then moving the folder directly
    var parent = dst.substring(0, dst.lastIndexOf("/"));
    await this.ensurePath(parent, !this.context || this.context.length < 2 ? 2 : 4);

    // move files, keep both may result in a renamed folder if the destination already exists
    var folder = await this.sp.web.getFolderByServerRelativePath(src).moveByPath(dst, {
      KeepBoth: true,
      RetainEditorAndModifiedOnMove: true,
      ShouldBypassSharedLocks: true
    });

    // return where its been moved
    return decodeURIComponent((await folder()).ServerRelativeUrl || destination);
  }

  public async saveFiles(path: string, additional: string | undefined, url: any | undefined, files: any, metadata: any | undefined): Promise<void> {
    if (path.indexOf("://") >= 0)
      path = path.substring(path.indexOf('/', 9));
    path = decodeURIComponent(path).replace(/\/$/, '');

    // common metadata for folder and each file, unless overridden at a file level
    var commonmeta = metadata ? JSON.parse(JSON.stringify(metadata)) : {};
    if (url)
      commonmeta['Request'] = url;

    var errors: Array<string> = [];
    try {
      var folder;
      if (metadata || url) {
        folder = await this.sp.web.getFolderByServerRelativePath(path).getItem();
        await folder.update(commonmeta);
      }

      // subfolders for these
      if (additional) {
        path += '/' + additional;
        folder = await this.sp.web.getFolderByServerRelativePath(path).getItem();
        await folder.update(commonmeta);
      }

      // process saves and deletes
      for (var i = 0; i < files.results.length; i++) {
        try {
          var file = files.results[i];
          if (!file.ListItemAllFields)
            file.ListItemAllFields = {};

          // clone common metadata for files
          for (var m in commonmeta) {
            file.ListItemAllFields[m] = commonmeta[m];
          }

          // basic list item fields cleanup
          delete file.ListItemAllFields["$$hashKey"];
          delete file.ListItemAllFields["Id"];
          delete file.ListItemAllFields["ID"];

          this.cleanSaveKeys(file.ListItemAllFields, file.OldListItemAllFields);

          if (file.Deleted) {
            // file to delete
            await this.sp.web.getFolderByServerRelativePath(path + '/' + file.FileName).recycle();
          } else if (file.Data) {
            // file to upload
            await this.sp.web.getFolderByServerRelativePath(path).files.addUsingPath(file.FileName, file.Data, { Overwrite: true });
            if (this.hasData(file.ListItemAllFields)) {
              let i = await this.sp.web.getFolderByServerRelativePath(path).files.getByUrl(file.FileName).getItem();
              await i.update(file.ListItemAllFields);
            }
            // mock the data back in so submit again doesnt fail
            file.TimeCreated = new Date();
            file.ServerRelativeUrl = path + '/' + file.FileName;
            delete file.Data;
          } else if (this.hasData(file.ListItemAllFields)) {
            // get current item and check for changes
            let i = await this.sp.web.getFolderByServerRelativePath(path + '/' + file.FileName).getItem();
            await i.update(file.ListItemAllFields);
          }
        } catch (e) {
          errors.push(`Error saving file ${file.FileName} in folder ${path} with error ${e}`);
        }
      }
    } catch (e) {
      alert('Error saving folder:\n\n' + e);
      throw e;
    }

    if (errors.length > 0) {
      alert('Error saving files:\n\n' + errors.join('\n'));
      throw errors.join('\n');
    }
  }
}
