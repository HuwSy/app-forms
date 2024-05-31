import { ErrorHandler } from '@angular/core';
import { App } from './App'

export class AngularLogging implements ErrorHandler {
    constructor() {
        // connection here
    }

    public async handleError(error: any): Promise<void> {
        console.error(App.AppName);
        console.error(App.Release);
        console.error(error);
    }
}
