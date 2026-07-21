import { Directive, ElementRef, Input, OnChanges, OnDestroy, Renderer2, SecurityContext } from "@angular/core";
import { DomSanitizer } from "@angular/platform-browser";

export interface SharepointChoiceRenderPayload {
  content: string | Node | null | undefined;
  hyperlink?: string;
  hyperlinkTarget?: string;
}

export type SharepointChoiceRenderContent = string | Node | SharepointChoiceRenderPayload | null | undefined;

@Directive({
  selector: "[spcRenderNode]",
  standalone: true,
})
export class SharepointChoiceRender implements OnChanges, OnDestroy {
  @Input("spcRenderNode") content: SharepointChoiceRenderContent;

  private removeRenderClickListener?: () => void;
  private removeClickListener?: () => void;
  private removeKeydownListener?: () => void;

  constructor(
    private host: ElementRef<HTMLElement>,
    private renderer: Renderer2,
    private sanitizer: DomSanitizer,
  ) {}

  ngOnDestroy(): void {
    this.clearFallbackLink();
  }

  ngOnChanges(): void {
    const el = this.host.nativeElement;
    const payload = this.asPayload(this.content);

    this.clearFallbackLink();
    this.clearRenderClickListener();
    this.renderer.removeStyle(el, "cursor");
    this.renderer.removeAttribute(el, "role");
    this.renderer.removeAttribute(el, "tabindex");

    while (el.firstChild) {
      el.removeChild(el.firstChild);
    }

    if (payload.content === null || payload.content === undefined) {
      return;
    }

    if (typeof payload.content === "string") {
      // Match Angular's default [innerHTML] sanitization behavior.
      el.innerHTML = this.sanitizer.sanitize(SecurityContext.HTML, payload.content) ?? "";
    } else {
      // Append the actual DOM node so any event listeners/etc remain intact.
      // Note: a Node can only exist in one place in the DOM at a time.
      this.renderer.appendChild(el, payload.content);
    }

    const hasInteractiveContent = this.hasInteractiveContent(el);

    if (hasInteractiveContent) {
      this.captureRenderedClicks(el);
    }

    if (payload.hyperlink && !hasInteractiveContent) {
      this.applyFallbackLink(el, payload.hyperlink, payload.hyperlinkTarget);
    }
  }

  private asPayload(content: SharepointChoiceRenderContent): SharepointChoiceRenderPayload {
    if (content && typeof content === "object" && !(content instanceof Node) && "content" in content) {
      return content;
    }

    return { content };
  }

  private hasInteractiveContent(root: HTMLElement): boolean {
    const interactiveSelector = [
      "a[href]",
      "button",
      "input",
      "select",
      "textarea",
      "summary",
      "[role='button']",
      "[role='link']",
      "[contenteditable='true']",
      "[tabindex]",
      "[onclick]",
    ].join(",");

    if (root.querySelector(interactiveSelector)) {
      return true;
    }

    for (const node of Array.from(root.querySelectorAll<HTMLElement>("*"))) {
      if (typeof node.onclick === "function") {
        return true;
      }
    }

    return false;
  }

  private captureRenderedClicks(el: HTMLElement): void {
    this.removeRenderClickListener = this.renderer.listen(el, "click", (event: MouseEvent) => {
      if (this.findInteractiveElement(event.target, el)) {
        event.stopPropagation();
      }
    });
  }

  private findInteractiveElement(target: EventTarget | null, boundary: HTMLElement): HTMLElement | null {
    let element = target instanceof HTMLElement ? target : target instanceof Node ? target.parentElement : null;

    while (element && element !== boundary) {
      if (this.isInteractiveElement(element)) {
        return element;
      }

      element = element.parentElement;
    }

    return null;
  }

  private isInteractiveElement(element: HTMLElement): boolean {
    if (
      element.matches(
        "a[href],button,input,select,textarea,summary,[role='button'],[role='link'],[contenteditable='true'],[tabindex],[onclick]",
      )
    ) {
      return true;
    }

    return typeof element.onclick === "function";
  }

  private applyFallbackLink(el: HTMLElement, hyperlink: string, hyperlinkTarget?: string): void {
    this.renderer.setStyle(el, "cursor", "pointer");
    this.renderer.setAttribute(el, "role", "link");
    this.renderer.setAttribute(el, "tabindex", "0");

    const navigate = () => {
      const anchor = this.renderer.createElement("a") as HTMLAnchorElement;
      anchor.href = hyperlink;
      anchor.target = hyperlinkTarget || "_self";
      if (anchor.target === "_blank") anchor.rel = "noopener";
      anchor.click();
    };

    this.removeClickListener = this.renderer.listen(el, "click", (event: MouseEvent) => {
      event.stopPropagation();

      if (event.defaultPrevented) {
        return;
      }

      navigate();
    });

    this.removeKeydownListener = this.renderer.listen(el, "keydown", (event: KeyboardEvent) => {
      if (event.key !== "Enter" && event.key !== " ") {
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      navigate();
    });
  }

  private clearRenderClickListener(): void {
    this.removeRenderClickListener?.();
    this.removeRenderClickListener = undefined;
  }

  private clearFallbackLink(): void {
    this.removeClickListener?.();
    this.removeClickListener = undefined;
    this.removeKeydownListener?.();
    this.removeKeydownListener = undefined;
  }
}
