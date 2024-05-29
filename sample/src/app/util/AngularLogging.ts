import { ErrorHandler } from '@angular/core';

export class AngularLogging implements ErrorHandler {
    constructor() {
        // connection here
    }

    public async handleError(error: any): Promise<void> {
        console.error(error);
    }
}
