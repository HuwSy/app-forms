import { NgModule, CUSTOM_ELEMENTS_SCHEMA, ErrorHandler } from '@angular/core';
import { SharepointChoiceComponent } from './sharepoint-choice.component';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { AngularLogging } from './AngularLogging';
import { NgxEditorModule } from 'ngx-editor';

@NgModule({
  declarations: [
    SharepointChoiceComponent
  ],
  imports: [
    BrowserModule,
    FormsModule,
    BrowserAnimationsModule,
    NgxEditorModule
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  providers: [{
    provide: ErrorHandler,
    useClass: AngularLogging
  }],
  exports: [
    SharepointChoiceComponent
  ]
})
export class SharepointChoiceModule { 
  constructor() {}
}
