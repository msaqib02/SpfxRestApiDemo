import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxRestApiDemo.module.scss';
import * as strings from 'spfxRestApiDemoStrings';
import { ISpfxRestApiDemoWebPartProps } from './ISpfxRestApiDemoWebPartProps';


import {
  SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ISPHttpClientOptions
} from '@microsoft/sp-http';
import { IODataUser } from '@microsoft/sp-odata-types';

export default class SpfxRestApiDemoWebPart extends BaseClientSideWebPart<ISpfxRestApiDemoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
    this.doSpfxGetRequest();
    this.doSpfxPostRequest();


  }
  private doSpfxGetRequest(): void {
    //current url can be accessed by from page context in spfx Webpart 
    // 'this' is our current Webpart object
    const hostUrl: string = this.context.pageContext.web.absoluteUrl;
    // create spHttpClient object to make Get/Post Request
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
    //Spfx GET call to access current user 
    spHttpClient.get(`${hostUrl}/_api/web/currentuser`, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      response.json().then((user: IODataUser) => {
        console.log(user.Title);
      });
    });

  }
  private doSpfxPostRequest(): void {
    //current url can be accessed by from page context in spfx Webpart 
    // 'this' is our current Webpart object
    const hostUrl: string = this.context.pageContext.web.absoluteUrl;
    // create spHttpClient object to make Get/Post Request
    const spHttpClient: SPHttpClient = this.context.spHttpClient;

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: `{ Title: 'RestPostedList', BaseTemplate: 100 }`
    };

    spHttpClient.post(`${hostUrl}/_api/web/lists`, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
         console.log(`${response.status}`);
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
