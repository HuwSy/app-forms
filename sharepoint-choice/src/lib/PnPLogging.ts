import { ILogListener } from "@pnp/logging";
import { App } from './App'

export class PnPLogging implements ILogListener {
    constructor() {
        // connection here
    }

    public async log(entry: any): Promise<void> {
        console.error(App.AppName);
        console.error(App.Release);
        console.error(entry);
    }
}
