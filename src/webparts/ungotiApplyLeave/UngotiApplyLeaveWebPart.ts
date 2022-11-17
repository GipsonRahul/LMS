import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,       
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'UngotiApplyLeaveWebPartStrings';
import UngotiApplyLeave from './components/UngotiApplyLeave';
import { IUngotiApplyLeaveProps } from './components/IUngotiApplyLeaveProps';

import { MSGraphClient, HttpClient } from '@microsoft/sp-http';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

export interface IUngotiApplyLeaveWebPartProps {
  description: string;
  card: boolean;
  list: boolean;
  cardTitle: string;
  listTitle: string;
  managerTitle:string;
  hrTitle:string;
  holidaysTitle:string;
  chkmanager: boolean;
  chkHR: boolean;
  UserRequest:boolean;
  chkHolidays:boolean;
  color: string;
}

export default class UngotiApplyLeaveWebPart extends BaseClientSideWebPart<IUngotiApplyLeaveWebPartProps> {

  public render(): void {
    var currentContext = this.context;
    this.context.msGraphClientFactory.getClient()
      .then((_graphClient: MSGraphClient): void => {
        const element: React.ReactElement<IUngotiApplyLeaveProps> = React.createElement(
          UngotiApplyLeave,
          {
            // description: this.properties.description,
            currentContext: currentContext,
            siteUrl: this.context.pageContext.web.absoluteUrl,
            card: this.properties.card,
            list: this.properties.list,
            cardTitle: this.properties.cardTitle,
            listTitle: this.properties.listTitle,
            managerTitle:this.properties.managerTitle,
            hrTitle:this.properties.hrTitle,
            holidaysTitle:this.properties.holidaysTitle,
            graphClient: _graphClient,
            chkmanager: this.properties.chkmanager,
            chkHR: this.properties.chkHR,
            UserRequest:this.properties.UserRequest,
            chkHolidays:this.properties.chkHolidays,
            color: this.properties.color
          }
        );
        ReactDom.render(element, this.domElement);
      });
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
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('cardTitle', {
                  label: 'Balance Title'
                }),
                PropertyPaneTextField('listTitle', {
                  label: 'List Title'
                }),
                PropertyPaneTextField('managerTitle', {
                  label: 'Manager Title'
                }),
                PropertyPaneTextField('hrTitle', {
                  label: 'HR Title'
                }),
                PropertyPaneTextField('holidaysTitle', {
                  label: 'Holidays Title'
                }),
              ]
            },
            {
              groupName: "Show these components",
              groupFields: [
                PropertyPaneCheckbox('card', {
                  checked: false,
                  text: "Balances"
                }),
                PropertyPaneCheckbox('list', {
                  checked: false,
                  text: "List"
                }),
                PropertyPaneCheckbox('chkmanager', {
                  checked: false,
                  text: "Manager"
                }),
                PropertyPaneCheckbox('chkHR', {
                  checked: false,
                  text: "HR"
                }),
                PropertyPaneCheckbox('UserRequest', {
                  checked: false,
                  text: "User & Request"
                }),
                PropertyPaneCheckbox('chkHolidays', {
                  checked: false,
                  text: "Holidays"
                }),
                PropertyFieldColorPicker('color', {
                  label: 'Color',
                  selectedColor: this.properties.color,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
