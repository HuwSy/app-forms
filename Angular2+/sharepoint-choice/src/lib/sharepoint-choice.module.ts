import { NgModule, CUSTOM_ELEMENTS_SCHEMA, ErrorHandler } from '@angular/core';
import { SharepointChoiceComponent } from './sharepoint-choice.component';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import { HttpClientModule } from '@angular/common/http';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { EditorModule } from '@tinymce/tinymce-angular';
import { AngularLogging } from './AngularLogging';

@NgModule({
  declarations: [
    SharepointChoiceComponent
  ],
  imports: [
    BrowserModule,
    FormsModule,
    HttpClientModule,
    BrowserAnimationsModule,
    EditorModule
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
