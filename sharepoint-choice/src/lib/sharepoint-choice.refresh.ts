import { Injectable } from "@angular/core";
import { Subject } from "rxjs";
import { SharepointChoiceForm, SharepointChoiceList } from "./sharepoint-choice.models";

export interface SharepointChoiceEvent {
  sourceId: number;
  form?: SharepointChoiceForm;
  spec?: SharepointChoiceList;
  field?: string;
}

@Injectable({ providedIn: "root" })
export class SharepointChoiceRefresh {
  private readonly refreshSubject = new Subject<SharepointChoiceEvent>();

  public readonly changes$ = this.refreshSubject.asObservable();

  public emit(event: SharepointChoiceEvent): void {
    this.refreshSubject.next(event);
  }

  /**
   * Triggers a refresh for any SharepointChoiceComponent instances bound to the same form reference.
   * Use this after mutating the form object outside of SharepointChoiceComponent.
   */
  public refreshForm(form: SharepointChoiceForm, field?: string): void {
    this.emit({ sourceId: 0, form, field });
  }

  /**
   * Triggers a refresh for any SharepointChoiceComponent instances bound to the same spec reference.
   * Use this after mutating the spec object outside of SharepointChoiceComponent.
   */
  public refreshSpec(spec: SharepointChoiceList, field?: string): void {
    this.emit({ sourceId: 0, spec, field });
  }

  /**
   * Convenience method to refresh by form and/or spec reference.
   */
  public refresh(options: { form?: SharepointChoiceForm; spec?: SharepointChoiceList; field?: string }): void {
    this.emit({ sourceId: 0, ...options });
  }
}
