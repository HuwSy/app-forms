import { ILogListener } from "@pnp/logging";
import pnp from '@pnp/pnpjs';
import { App } from './App'

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