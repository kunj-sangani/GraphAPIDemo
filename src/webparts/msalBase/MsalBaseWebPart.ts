import * as React from 'react';
import * as Msal from "msal";
import * as ReactDom from 'react-dom';
import { MSGraphClient } from "@microsoft/sp-http";
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartContext
} from '@microsoft/sp-webpart-base';

import * as strings from 'MsalBaseWebPartStrings';
import MsalBase from './components/MsalBase';
import { IMsalBaseProps } from './components/IMsalBaseProps';

export interface IMsalBaseWebPartProps {
  description: string;
  msalObjcet: Msal.UserAgentApplication;
  context: WebPartContext;
}

const msalConfig: Msal.Configuration = {
  auth: {
    clientId: 'cc24ab80-701d-46b2-88fe-8b4f2cf77cfc',
    redirectUri: "https://testinglala.sharepoint.com/SitePages/Graph-API-Expamples.aspx"
  }, cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true
  }
};
const msalInstance = new Msal.UserAgentApplication(msalConfig);

export default class MsalBaseWebPart extends BaseClientSideWebPart<IMsalBaseWebPartProps> {

  public render(): void {
    let name: string = null;
    msalInstance.handleRedirectCallback((error, response) => {
      name = response.account.userName;
    });
    let loginRequest = {
      scopes: ["user.read"]
    };
    if (msalInstance.getAccount() || name) {
      const element: React.ReactElement<IMsalBaseProps> = React.createElement(
        MsalBase,
        {
          description: this.properties.description,
          msalObjcet: msalInstance,
          context: this.context
        }
      );
      ReactDom.render(element, this.domElement);
    } else {
      msalInstance.loginRedirect(loginRequest);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
