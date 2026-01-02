import { Directive, ElementRef, Input, OnChanges, Renderer2, SecurityContext } from '@angular/core';
import { DomSanitizer } from '@angular/platform-browser';

export type SharepointChoiceRenderContent = string | Node | null | undefined;

@Directive({
  selector: '[spcRenderNode]',
  standalone: true
})
export class SharepointChoiceRender implements OnChanges {
  @Input('spcRenderNode') content: SharepointChoiceRenderContent;

  constructor(
    private host: ElementRef<HTMLElement>,
    private renderer: Renderer2,
    private sanitizer: DomSanitizer
  ) { }

  ngOnChanges(): void {
    const el = this.host.nativeElement;

    while (el.firstChild) {
      el.removeChild(el.firstChild);
    }

    if (this.content === null || this.content === undefined) {
      return;
    }

    if (typeof this.content === 'string') {
      // Match Angular's default [innerHTML] sanitization behavior.
      el.innerHTML = this.sanitizer.sanitize(SecurityContext.HTML, this.content) ?? '';
      return;
    }

    // Append the actual DOM node so any event listeners/etc remain intact.
    // Note: a Node can only exist in one place in the DOM at a time.
    this.renderer.appendChild(el, this.content);
  }
}
