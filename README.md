# app-forms

A simple Angular framework for rapid development originally in AngularJS/JavaScript and progressed to Angular18+ within an SPFx wrapper giving controls which can use this.spec to determin the field details and manipulate this.form, the sample is based on the older Module version <= 6 whereas newer >= 7 is made as an angular Component as per the details and usage below

Requires node 20, nvm can be used effectively as can project leve installed angular

```
npm install @angular/cli@18
```
Add to package.json
```
  "scripts": {
    "new": "del package* && ng new --commit=false --routing=false --style=scss --directory .\\"
  },
```
Then
```
npm run new <solution>
```
remove default app.component.* app.config.* styles.scss and registrations in main.ts and app.module.ts  

Component only apps have issues within SharePoint page navigation and SPFx variable injections therefore using NgModule still

src/main.ts
```
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app/app.module';
platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.error(err));
```

src/app/app.module.ts
```
import { NgModule, Injector, ErrorHandler } from '@angular/core';
import { createCustomElement } from '@angular/elements';
import { FormsModule } from '@angular/forms';
import { BrowserModule } from '@angular/platform-browser';

import { SharepointChoiceComponent } from '@qicglobal/sharepoint-choice';
import { AngularLogging } from '../../App';

import { AgGridAngular } from 'ag-grid-angular';
import { AllCommunityModule, ModuleRegistry } from 'ag-grid-community'; 
ModuleRegistry.registerModules([AllCommunityModule]);

// repeat this per web part
import { <webpart>Component } from './<webpart>/<webpart>.component';

@NgModule({
  declarations: [
    // repeat this per web part
    <webpart>Component
  ],
  imports: [
    BrowserModule,
    FormsModule,
    SharepointChoiceComponent,
    AgGridAngular
  ],
  providers: [{
    provide: ErrorHandler,
    useClass: AngularLogging
  }]
})
export class AppModule {
  constructor(private injector: Injector) { }

  ngDoBootstrap() {
    // repeat this per web part
    customElements.define('app-<webpart>', createCustomElement(<webpart>Component, { injector: this.injector }));
  }
}
```

New web parts / components / dashboards / forms etc
```
npm run ng generate component --style=scss <webpart>
```

To be added to html templates, repeat app-choice as required
```
<form ngNativeValidate #input (keydown.enter)="enterKey($event)">
  <app-choice [form]="form"
              [field]="'Title'"
              [spec]="spec">
  </app-choice>
</form>
```

To be added to all web parts
```
import { SharepointChoiceUtils } from 'sharepoint-choice';
```
```
    // allow spfx property to override web used or where related pages exist or any other app related deployed params etc
    @Input() context!: string;

    // register the utils
    this.util = new SharepointChoiceUtils(this.context ?? null);

    // load user and permission details
    this.util.permissions().then(r => {
        this.userId = r.userId;
        this.perm = r.perms;
    });

    // load the list field spec
    this.util.fields(this.list).then(r => {
        this.spec = r;
    });

    // load the list item
    this.util.data(id, this.list).then(d => {
        this.form = d;
        this.uned = JSON.parse(JSON.stringify(this.form));
    });

    // load version history
    this.util.history(id, this.list).then(d => {
        this.versions = d
    });

    // save the form and any attachments on the form object
    this.util.save(this.form, this.uned, this.list).then(id => {
        this.form.Id = id;
    });

    // load an api, or post to an api
    this.util.callApi(
        `guid`,
        `permission`,
        'path',
        App.APIRelease || App.Release,
        'POST',
        {'content':'dummy'}).then(results => {
        this.results = results;
    });
```
```
  enterKey(e:Event|any):void {
    // |any exists is here because Event srcElement is deprecated
    if (e.srcElement.tagName != 'TEXTAREA')
      e.preventDefault();
  }
```

Make the index.html multi webpart test compatible, remove app-root and replace with
```
  <script>
    let e = 'root';
    if (document.location.search != '')
      e = document.location.search.replace('?','');
    let c = document.createElement(`app-${e}`)
    document.body.appendChild(c);
  </script>
```

To be added to tsconfig.json to avoid some errors with pnp
```
  "skipLibCheck": true,
```

To add to package.json for easier use of the wrapper
```
  "bundle": "ng build --aot --delete-output-path --output-hashing=none"
```

angular.json build options if using esbuild/application
```
  "index": { "input": "src/index.html", "preloadInitial": false },
```

angular.json other useful changes
styles
```
  "node_modules/sharepoint-choice/src/styles.scss"
```
"type": "initial"
```
  "maximumError": "5mb"
```
"type": "anyComponentStyle"
```
  "maximumError": "512kb"
```

.vscode/launch.json
```
  "url": "https://<tennant>.sharepoint.com/sites/<site name>"
```

To generate SSL certs for debugging within localhost and SPFx wrapper
```
npm install -g office-addin-dev-certs
office-addin-dev-certs install --days 3650
```
Which will need this to be added to angular.json "serve":
```
  "development": {
    "browserTarget": "<solution>:build:development",
    "publicHost": "localhost",
    "port": 443,
    "ssl": true
  }
},
"options": {
  "allowedHosts": [
    "localhost",
    "<tennant>.sharepoint.com"
  ],
  "sslKey": "./localhost.key",
  "sslCert": "./localhost.crt"
}
```
