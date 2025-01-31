import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';

export interface IAngularWrapperWebPartProps {
  tag: string;
  src: string;
  adt: string;
  esb: boolean;
}

export default class AngularWrapperWebPart extends BaseClientSideWebPart<IAngularWrapperWebPartProps> {
  public render(): void {
    if ((this.properties.src || '') === '' || (this.properties.tag || '') === '') {
      this.domElement.innerHTML = `<b>URL and TAG properties must be entered</b>`;
      return;
    }

    let src = this.properties.src;
    if (src.toLowerCase().indexOf('https://') === 0) {
      // absolute url
      src = src.substring(src.indexOf('/', 9));
    } else if (src.indexOf('/') === 0) {
      // server relative url
      src = window.origin + src;
    } else {
      // web relative url
      src = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '') + '/' + src;
    }

    if (this.properties.esb) {
      if (~src.toLowerCase().indexOf('://localhost'))
        this.requireClientSide(src + '/@vite/client');
      
      this.requireClientSide(src + '/styles.css');
      this.requireClientSide(src + '/polyfills.js');
      this.requireClientSide(src + '/main.js');
    } else {
      this.requireClientSide(src + '/polyfills.js');
      this.requireClientSide(src + '/runtime.js');
      this.requireClientSide(src + '/main.js');
      this.requireClientSide(src + '/styles.css');
      
      if (~src.toLowerCase().indexOf('//localhost'))
        this.requireClientSide(src + '/vendor.js');
    }

    const tag = this.properties.tag.replace(/^<\//, '').replace(/^</, '').replace(/>$/, '');
    this.domElement.innerHTML = `
      <${ tag } ${ (this.properties.adt || '').replace(/>/g, '') } context="${this.context.pageContext.web.absoluteUrl}">Loading app...</${ tag }>
    `;

    // suppress gulp serve warning when using this on workbench
    setTimeout(() => {
      const b = document.getElementsByTagName('button');
      for (let i = 0; i < b.length; i++)
          if (b[i].getAttribute('data-automation-id') === "GulpServeWarningOkButton")
              b[i].click();
    },500);
  }

  private requireClientSide(file:string):void {
    let p;
    if (~file.indexOf('.css')) {
      p = document.createElement('link');
      p.rel = 'stylesheet';
      p.href = file;
    } else {
      p = document.createElement('script');
      p.src = file;
      if (this.properties.esb) p.type = 'module';
    }
    document.head.append(p);
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "SPFx Angular Wrapper"
          },
          groups: [
            {
              groupName: "SPFx settings",
              groupFields: [
                PropertyPaneTextField('src', {
                  label: "Folder",
                  description: "URL to script folder, i.e. SiteAssets/app"
                }),
                PropertyPaneTextField('tag', {
                  label: "Selector",
                  description: "Angular selector TAG name, i.e. app-tag"
                }),
                PropertyPaneTextField('adt', {
                  label: "Additional",
                  description: "Additional attributes, i.e. data='test'"
                }),
                PropertyPaneCheckbox('esb', {
                  text: "ESBuild used",
                  checked: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
