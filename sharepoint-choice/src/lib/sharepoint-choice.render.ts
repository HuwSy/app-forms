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

  public constructor(
    private host: ElementRef<HTMLElement>,
    private renderer: Renderer2,
    private sanitizer: DomSanitizer,
  ) {}

  public ngOnDestroy(): void {
    this.clearFallbackLink();
  }

  public ngOnChanges(): void {
    const hostElement = this.host.nativeElement;
    const content = this.content;
    const renderPayload: SharepointChoiceRenderPayload =
      content && typeof content === "object" && !(content instanceof Node) && "content" in content
        ? content
        : { content };

    this.clearFallbackLink();
    this.clearRenderClickListener();
    this.renderer.removeStyle(hostElement, "cursor");
    this.renderer.removeAttribute(hostElement, "role");
    this.renderer.removeAttribute(hostElement, "tabindex");

    while (hostElement.firstChild) {
      hostElement.removeChild(hostElement.firstChild);
    }

    if (renderPayload.content === null || renderPayload.content === undefined) {
      return;
    }

    if (typeof renderPayload.content === "string") {
      // Match Angular's default [innerHTML] sanitization behavior.
      hostElement.innerHTML = this.sanitizer.sanitize(SecurityContext.HTML, renderPayload.content) ?? "";
    } else {
      // Append the actual DOM node so any event listeners/etc remain intact.
      // Note: a Node can only exist in one place in the DOM at a time.
      this.renderer.appendChild(hostElement, renderPayload.content);
    }

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

    let hasInteractiveContent = hostElement.querySelector(interactiveSelector) !== null;
    if (!hasInteractiveContent) {
      for (const node of Array.from(hostElement.querySelectorAll<HTMLElement>("*"))) {
        if (typeof node.onclick === "function") {
          hasInteractiveContent = true;
          break;
        }
      }
    }

    if (hasInteractiveContent) {
      this.captureRenderedClicks(hostElement);
    }

    if (renderPayload.hyperlink && !hasInteractiveContent) {
      this.applyFallbackLink(hostElement, renderPayload.hyperlink, renderPayload.hyperlinkTarget);
    }
  }

  private captureRenderedClicks(el: HTMLElement): void {
    this.removeRenderClickListener = this.renderer.listen(el, "click", (event: MouseEvent) => {
      for (
        let currentElement =
          event.target instanceof HTMLElement
            ? event.target
            : event.target instanceof Node
              ? event.target.parentElement
              : null;
        currentElement && currentElement !== el;
        currentElement = currentElement.parentElement
      ) {
        if (
          currentElement.matches(
            "a[href],button,input,select,textarea,summary,[role='button'],[role='link'],[contenteditable='true'],[tabindex],[onclick]",
          ) ||
          typeof currentElement.onclick === "function"
        ) {
          event.stopPropagation();
          break;
        }
      }
    });
  }

  private applyFallbackLink(el: HTMLElement, hyperlink: string, hyperlinkTarget?: string): void {
    this.renderer.setStyle(el, "cursor", "pointer");
    this.renderer.setAttribute(el, "role", "link");
    this.renderer.setAttribute(el, "tabindex", "0");

    const navigate = () => {
      const anchor = this.renderer.createElement("a") as HTMLAnchorElement;
      // This callback URL intentionally bypasses Angular's template URL sanitizer to preserve existing consumer behavior.
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
