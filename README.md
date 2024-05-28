# app-forms

A simple Angular framework for rapid development originally in AngularJS/JavaScript and progressed to Angular18 within an SPFx wrapper giving controls which can use this.spec to determin the field details and manipulate this.form

To be added to app module
```
import { SharepointChoiceModule } from 'sharepoint-choice';
```
```
  imports: [
    BrowserModule,
    FormsModule,
    HttpClientModule,
    BrowserAnimationsModule,
    SharepointChoiceModule
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
```

To be added to html templates
```
<app-choice [form]="form" [spec]="spec" [field]="'Title'"></app-choice>
```

To be added to the components
```
import { SharepointChoiceUtils } from 'sharepoint-choice';
```
```
// register the utils
this.util = new SharepointChoiceUtils(this.context);

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

// save the form
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

Dependencies
```
  "dependencies": {
    "@angular/animations": "^18.0.0",
    "@angular/common": "^18.0.0",
    "@angular/compiler": "^18.0.0",
    "@angular/core": "^18.0.0",
    "@angular/elements": "^18.0.0",
    "@angular/forms": "^18.0.0",
    "@angular/platform-browser": "^18.0.0",
    "@angular/platform-browser-dynamic": "^18.0.0",
    "@angular/router": "^18.0.0",
    "@pnp/pnpjs": "^2.15.0",
    "sharepoint-choice": "^4.0.0",
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

To be added to tsconfig.json to avoid some errors
```
  "strict": false,
  "skipLibCheck": true,
```

To add to package.json for easier use of the wrapper
```
  "bundle": "ng build --aot --build-optimizer --delete-output-path --output-hashing none"
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
    "ssl": true,
    "disableHostCheck": true
  }
},
"options": {
  "sslKey": "./localhost.key",
  "sslCert": "./localhost.crt"
}
```
