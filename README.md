# app-forms

A simple Angular framework for rapid development originally in AngularJS/JavaScript and progressed to Angular18+ within an SPFx wrapper giving controls which can use this.spec to determin the field details and manipulate this.form, the sample is based on the older Model version 6 whereas newer 7+ is Component as per the details below

Requires node 20, nvm can be used effectively as can project installed angular
```
npm install @angular/cli@18
```

Dependencies
```
  "dependencies": {
    "@angular/animations": "^18.0.0",
    "@angular/common": "^18.0.0",
    "@angular/compiler": "^18.0.0",
    "@angular/core": "^18.0.0",
    "@angular/forms": "^18.0.0",
    "@angular/platform-browser": "^18.0.0",
    "@angular/platform-browser-dynamic": "^18.0.0",
    "@angular/router": "^18.0.0",
    "@pnp/pnpjs": "^2.15.0",
    "rxjs": "~7.8.1",
    "tslib": "^2.6.2",
    "zone.js": "~0.14.3"
  },
  "devDependencies": {
    "@angular-devkit/build-angular": "^18.0.1",
    "@angular/cli": "^18.0.1",
    "@angular/compiler-cli": "^18.0.0",
    "@types/jasmine": "~5.1.0",
    "jasmine-core": "~5.1.0",
    "typescript": "~5.4.2"
  }
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
remove default app.component.* app.config.* styles.scss, app.module.ts and registrations in main.ts

New web parts/components/dashboards/forms etc
```
npm run ng generate component --style=scss <webpart>
```

src/main.ts
```
import { bootstrapApplication } from '@angular/platform-browser';

// repeat per component
import { <webpart>Component } from './app/<webpart>/<webpart>.component';
bootstrapApplication(<webpart>Component)
  .catch((err) => console.error(err));
```

To be added to html templates
```
<form ngNativeValidate #input (keydown.enter)="enterKey($event)">
  <app-choice [form]="form"
              [field]="'Title'"
              [spec]="spec">
  </app-choice>
</form>
```

To be added to the components
```
import { SharepointChoiceComponent, SharepointChoiceUtils } from 'sharepoint-choice';
import { AngularLogging } from './App';
```
```
  standalone: true,
  imports: [
    CommonModule,             // ngIf, ngFor etc
    FormsModule,              // Forms on screen
    SharepointChoiceComponent,// App-Choice fields
  ],
  providers: [{
    provide: ErrorHandler,
    useClass: AngularLogging
  }]
```
```
    // allow spfx property to override web used
    @Input() context!: string;
    // register the utils
    this.util = new SharepointChoiceUtils(this.context || null);

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

Make the index.html multi webpart, remove app-root and replace with
```
  <script>
    let e = 'root';
    if (document.location.search != '')
      e = document.location.search.replace('?','');
    let c = document.createElement(`app-${e}`)
    document.body.appendChild(c);
  </script>
```

To be added to tsconfig.json to avoid some errors
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
