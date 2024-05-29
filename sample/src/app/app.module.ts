import { BrowserModule } from '@angular/platform-browser';
import { NgModule, Injector, CUSTOM_ELEMENTS_SCHEMA, ErrorHandler } from '@angular/core';
import { createCustomElement } from '@angular/elements';
import { SharepointChoiceModule } from 'sharepoint-choice';
import { FormsModule } from '@angular/forms';
import { HttpClientModule } from '@angular/common/http';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { AngularLogging } from './util/AngularLogging';

import { HelloWorldWebPartComponent } from './hello-world-web-part/hello-world-web-part.component';

@NgModule({
  declarations: [
    HelloWorldWebPartComponent
  ],
  imports: [
    BrowserModule,
    FormsModule,
    HttpClientModule,
    BrowserAnimationsModule,
    SharepointChoiceModule
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  providers: [{
    provide: ErrorHandler,
    useClass: AngularLogging
  }],
  entryComponents: [HelloWorldWebPartComponent]
})
export class AppModule {
  constructor(private injector: Injector) {}

  ngDoBootstrap() {
    const el = createCustomElement(HelloWorldWebPartComponent, { injector: this.injector });
    customElements.define('app-hello-world-web-part', el);
  }
}
