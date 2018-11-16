import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'CalendarMgrWebPartStrings';
import CalendarMgr from './components/CalendarMgr';
import { ICalendarMgrProps } from './components/ICalendarMgrProps';

export interface ICalendarMgrWebPartProps {
  groupId: String;
}

export default class CalendarMgrWebPart extends BaseClientSideWebPart<ICalendarMgrWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICalendarMgrProps > = React.createElement(
      CalendarMgr,
      {
        groupId: "8d6dcaf7-8835-4bc7-9b6f-91a52972e4af",
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
                PropertyPaneTextField('groupId', {
                  label: "Group Id"
                }),      
              ]
            }
          ]
        }
      ]
    };
  }
}
