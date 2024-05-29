import { ILogListener } from "@pnp/logging";
import { ErrorHandler } from '@angular/core';
import pnp from '@pnp/pnpjs';

export const App = {
    AppName: 'Sample',
    Category: 'SPO',
    Release: ~document.location.href.toLowerCase().indexOf('/prd-') ? 'LIVE'
        : ~document.location.href.toLowerCase().indexOf('/pre-') ? 'PRE'
        : ~document.location.href.toLowerCase().indexOf('/tst-') ? 'TST'
        : ~document.location.href.toLowerCase().indexOf('/sit-') ? 'SIT'
        : ~document.location.href.toLowerCase().indexOf('localhost') ? 'LOCAL'
        : 'DEV',
    AzureApp: '',
    // unset from null to allow override of API to query
    APIRelease: null
}

export class AngularLogging implements ErrorHandler {
    private _user: string;

    constructor() {
        setTimeout(() => { pnp.sp.web.currentUser.get().then((u) => this._user = (u.LoginName), () => {}); }, 100);
    }

    public async handleError(error: any): Promise<void> {
        console.log(error);
    }
}

export class PnPLogging implements ILogListener {
    private _user: string;

    constructor() {
        setTimeout(() => { pnp.sp.web.currentUser.get().then((u) => this._user = (u.LoginName), () => {}); }, 100);
    }

    public async log(entry: any): Promise<void> {
        console.log(entry);
    }
}