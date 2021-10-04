Mock of app-forms/Common/choices.js in SPFx and Angular2+ parked here in case its of use. Recreating the whole app-forms in SPFx, Angular2+ and PNPJS was less useful overall as PNPJS managed/executed cleanly enough to not need a framework like this in most cases.

Assumes in app.module.ts

import { FormsModule } from '@angular/forms'; 
imports: [...,FormsModule]
