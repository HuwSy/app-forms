import { bootstrapApplication } from '@angular/platform-browser';
import { SharepointChoiceLogging } from 'sharepoint-choice';
import { ErrorHandler, provideZonelessChangeDetection } from '@angular/core';
import { SampleComponent } from './app/sample/sample';

// only bootstrap required components
var loadComponents = () => {
  let loading = false;
  [{
    com: SampleComponent,
    tag: 'app-sample'
  }].map((component) => {
    var el = document.querySelector(component.tag);
    // no element or already loaded then skip
    if (!el || el.hasAttribute('loaded'))
      return;
    loading = true;
    // flag loaded on this page
    el.setAttribute('loaded', 'true');
    // bootstrap the component
    bootstrapApplication(component.com, {
        providers: [provideZonelessChangeDetection(), {provide: ErrorHandler, useClass: SharepointChoiceLogging}]
      })
      .catch((err) => console.error(err));
  });
  if (!loading)
    setTimeout(loadComponents, 500);
};

// on partial page load trigger bootstrap load
setTimeout(() => {
  var w = window as any;
  w.pushStateOriginal = w.history.pushState.bind(w.history);
  w.history.pushState = function () {
    w.pushStateOriginal(...Array.prototype.slice.call(arguments, 0));
    setTimeout(loadComponents, 500);
  };
  w.addEventListener('popstate', () => {
    setTimeout(loadComponents, 500)
  });
}, 500);

// on page load trigger bootstrap load
loadComponents();
