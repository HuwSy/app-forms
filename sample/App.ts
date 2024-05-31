import { ErrorHandler } from '@angular/core';

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
    constructor() {
    }

    public async handleError(error: any): Promise<void> {
        console.log(error);
    }
}