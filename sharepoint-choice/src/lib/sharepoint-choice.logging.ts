import { ErrorHandler } from '@angular/core';
import { fromError } from 'stacktrace-js';
import { App } from './App'

export class SharepointChoiceLogging implements ErrorHandler {
    constructor() {
        // connection here
    }

    public async handleError(error: any): Promise<void> {
        let w:any = window;
        if (w._spPageContextInfo) {
            console.error(w._spPageContextInfo.userLoginName);
            console.error(w._spPageContextInfo.webAbsoluteUrl);
            console.error(w._spPageContextInfo.webTitle);
        }
        console.error(App.Release);
        console.error(document.location.href);
        console.error(document.location.search);
        console.error(document.location.hash);
        try {
            console.error(JSON.stringify(await fromError(entry, {offline: true})));
        } catch (e) {
            console.error(JSON.stringify(entry.data ? entry.data.StackTrace : null));
        }
    }
}
