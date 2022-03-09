import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'UserSummaryDashboardWebPartStrings';
import UserSummaryDashboard from './components/UserSummaryDashboard';
import { IUserSummaryDashboardProps } from './components/IUserSummaryDashboardProps';

export interface IUserSummaryDashboardWebPartProps {
  description: string;
  leaveLink: string;
  attendanceLink: string;
}

export default class UserSummaryDashboardWebPart extends BaseClientSideWebPart<IUserSummaryDashboardWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IUserSummaryDashboardProps> = React.createElement(
      UserSummaryDashboard,
      {
        description: this.properties.description,
        leaveLink: this.properties.leaveLink,
        attendanceLink: this.properties.attendanceLink,
        context: this.context
      }
    );

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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('leaveLink', {
                  label: "Leave Link",
                }),
                PropertyPaneTextField('attendanceLink', {
                  label: "Attendance Link",
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
