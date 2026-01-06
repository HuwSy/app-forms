import { Injectable } from '@angular/core';
import { Subject } from 'rxjs';
import { SharepointChoiceForm, SharepointChoiceList } from './sharepoint-choice.models';

export interface SharepointChoiceEvent {
  sourceId: number;
  form?: SharepointChoiceForm;
  spec?: SharepointChoiceList;
  field?: string;
}

@Injectable({ providedIn: 'root' })
export class SharepointChoiceRefresh {
  private readonly subject = new Subject<SharepointChoiceEvent>();

  readonly changes$ = this.subject.asObservable();

  emit(event: SharepointChoiceEvent): void {
    this.subject.next(event);
  }

  /**
   * Triggers a refresh for any SharepointChoiceComponent instances bound to the same form reference.
   * Use this after mutating the form object outside of SharepointChoiceComponent.
   */
  refreshForm(form: SharepointChoiceForm, field?: string): void {
    this.emit({ sourceId: 0, form, field });
  }

  /**
   * Triggers a refresh for any SharepointChoiceComponent instances bound to the same spec reference.
   * Use this after mutating the spec object outside of SharepointChoiceComponent.
   */
  refreshSpec(spec: SharepointChoiceList, field?: string): void {
    this.emit({ sourceId: 0, spec, field });
  }

  /**
   * Convenience method to refresh by form and/or spec reference.
   */
  refresh(options: { form?: SharepointChoiceForm; spec?: SharepointChoiceList; field?: string }): void {
    this.emit({ sourceId: 0, ...options });
  }
}
