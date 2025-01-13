import { BrowserModule } from '@angular/platform-browser';
import { NgModule, Injector, ErrorHandler } from '@angular/core';
import { createCustomElement } from '@angular/elements';
import { SharepointChoiceComponent } from 'sharepoint-choice';
import { FormsModule } from '@angular/forms';
import { App, AngularLogging } from '../../App';

import { HelloWorldWebPartComponent } from './hello-world-web-part/hello-world-web-part.component';

@NgModule({
  declarations: [
    HelloWorldWebPartComponent
  ],
  imports: [
    BrowserModule,
    FormsModule,
    SharepointChoiceComponent
  ],
  providers: [{
    provide: ErrorHandler,
    useClass: AngularLogging
  }]
})
export class AppModule {
  constructor(private injector: Injector) {}

  ngDoBootstrap() {
    customElements.define('app-hello-world-web-part', createCustomElement(HelloWorldWebPartComponent, { injector: this.injector }));
  }
}
