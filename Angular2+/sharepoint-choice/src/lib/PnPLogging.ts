import { ILogListener } from "@pnp/logging";

export class PnPLogging implements ILogListener {
    constructor() {
        // connection here
    }

    public async log(entry: any): Promise<void> {
        console.error(entry);
    }
}
