# app-forms

A simple Angular framework for rapid development originally in AngularJS/JavaScript and progressed to currently run with Angular2+ within an SPFx wrapper giving controls which can use this.spec to determin the field details from the list schema loaded on page load and manipulate this.form for submitting into the list.

Project level installed angular below can be used to avoid additional directories and installs.

```
npm install @angular/cli@22
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

To generate SSL certs for debugging within localhost and SPFx wrapper
```
npm install -g office-addin-dev-certs
office-addin-dev-certs install --days 3650
```
Then copy these certs into the application directory.

