import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import CustomerSatisfaction from './components/CustomerSatisfaction';
import { ICustomerSatisfactionProps } from './components/ICustomerSatisfactionProps';

export interface ICustomerSatisfactionWebPartProps {
  teamLeaderListName: string;
}

export default class CustomerSatisfactionWebPart
  extends BaseClientSideWebPart<ICustomerSatisfactionWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICustomerSatisfactionProps> =
      React.createElement(CustomerSatisfaction, {
        context: this.context,
        currentUserEmail: this.context.pageContext.user.email,
        currentUserLoginName: this.context.pageContext.user.loginName,
        currentUserDisplayName: this.context.pageContext.user.displayName,
        teamLeaderListName: this.properties.teamLeaderListName || 'TeamLeader'
      });

    ReactDom.render(element, this.domElement);
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
          header: { description: 'Customer Satisfaction Program Settings' },
          groups: [
            {
              groupName: 'List Configuration',
              groupFields: [
                PropertyPaneTextField('teamLeaderListName', {
                  label: 'TeamLeader List Name'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}