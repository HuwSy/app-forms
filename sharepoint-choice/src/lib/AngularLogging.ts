import { ErrorHandler } from '@angular/core';
import pnp from '@pnp/pnpjs';
import { App } from './App'

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