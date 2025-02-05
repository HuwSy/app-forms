import { ILogListener } from "@pnp/logging";
import { App } from './App'

export class PnPLogging implements ILogListener {
    constructor() {
        // connection here
    }

    public async log(entry: any): Promise<void> {
        let w:any = window;
        console.error(w._spPageContextInfo.userLoginName);
        console.error(w._spPageContextInfo.webAbsoluteUrl);
        console.error(w._spPageContextInfo.webTitle);
        console.error(App.Release);
        console.error(document.location.search);
        console.error(document.location.hash);
        try {
            console.error(JSON.stringify(await fromError(entry, {offline: true})));
        } catch (e) {
            console.error(JSON.stringify(entry.data ? entry.data.StackTrace : null));
        }
    }
}
