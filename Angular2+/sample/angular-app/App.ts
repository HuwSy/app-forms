import { ILogListener } from "@pnp/logging";
import { ErrorHandler } from '@angular/core';
import pnp from '@pnp/pnpjs';

export const App = {
    AppName: '...',
    Release: ~document.location.href.toLowerCase().indexOf('/prd-') ? 'LIVE' 
        : ~document.location.href.toLowerCase().indexOf('/pre-') ? 'PRE' 
        : ~document.location.href.toLowerCase().indexOf('/tst-') ? 'TST' 
        : ~document.location.href.toLowerCase().indexOf('/sit-') ? 'SIT' 
        : 'DEV',
    // unset from null to allow override of API to query
    APIRelease: null
}

export class AngularLogging implements ErrorHandler {
    private user:string;

    constructor() {
        // connection here
        setTimeout(() => { pnp.sp.web.currentUser.get().then((u) => this.user = u.LoginName, () => {}); }, 100);
    }

    public async handleError(error: any): Promise<void> {
        console.error(this.user);
        console.error(App.AppName);
        console.error(App.Release);
        console.error(error);
    }
}

export class PnPLogging implements ILogListener {
    private user:string;

    constructor() {
        // connection here
        setTimeout(() => { pnp.sp.web.currentUser.get().then((u) => this.user = u.LoginName, () => {}); }, 100);
    }

    public async log(entry: any): Promise<void> {
        console.error(this.user);
        console.error(App.AppName);
        console.error(App.Release);
        console.error(entry);
    }
}