import { bootstrapApplication } from '@angular/platform-browser';

import { HelloWorldWebPartComponent } from './app/hello-world-web-part/hello-world-web-part.component';

// only bootstrap required components
var loadComponents = () => {
  [{
    com: HelloWorldWebPartComponent,
    tag: 'app-hello-world-web-part'
  }].forEach((component) => {
    var el = document.querySelector(component.tag);
    // no element or already loaded then skip
    if (!el || el.hasAttribute('loaded'))
      return;
    // flag loaded on this page
    el.setAttribute('loaded', 'true');
    // bootstrap the component
    bootstrapApplication(component.com)
      .catch((err) => console.error(err));
  });
};

// on partial page load trigger bootstrap load
setTimeout(() => {
  ['navigate','currententrychange'].forEach(evt => 
    window['navigation'].addEventListener(evt, () => {
      setTimeout(loadComponents, 2000)
    }, false)
  );
}, 2000);

// on page load trigger bootstrap load
loadComponents();
