import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {   } from "@microsoft/sp-http";


import * as strings from 'NewPartWebPartStrings';
//import {SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import NewPart from './components/NewPart';
import { ITestProps } from './components/INewPartProps';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
//import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

//import styles from './components/NewPart.module.scss';

export interface IPropertyControlsTestWebPartProps {
  lists: string | string[]; // Stores the list ID(s)
}
export interface INewPartWebPartProps {
  description: string;
  test2: string;
  lists: string;
 }

/* type PropsTypes = {
  Links: {
    Url: string
  }
  Title: string
} */


export default class NewPartWebPart extends BaseClientSideWebPart<INewPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
       
   /*  private async _getListData(): Promise<any> {
        const url1 = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Useful links')/items`;
        const url2 = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Useful links2')/items`;
        //const url3 = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Useful links2')/items({item_id})`;
       
        let url: string;
        this.properties.test2 === '1' ? url = url1 : url = url2;
        return await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
            return response.json();
          })
        .catch(() => {});
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
       this._renderList(response.value);
      })
      .catch(() => {});
  }
  
 */
  
  
 
  public render(): void {
    const element: React.ReactElement<ITestProps> = React.createElement(
      NewPart,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        selectedList: this.properties.lists,
        context: this.context
      }
     );
   
    ReactDom.render(element, this.domElement);
    /* this._renderListAsync(); */
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
                label: 'Description'
              }),
              PropertyPaneTextField('test', {
                label: 'Multi-line Text Field',
                multiline: true
              }),
              PropertyPaneCheckbox('test1', {
                text: 'Edit post'
              }),
              PropertyPaneDropdown('test2', {
                label: 'Select list',
                options: [
                  { key: '1', text: 'Useful links' },
                  { key: '2', text: 'Useful links2' },
                  
                ]}),
              PropertyPaneToggle('test3', {
                label: 'Toggle',
                onText: 'On',
                offText: 'Off'
              }),
              PropertyFieldListPicker('lists', {
                label: 'Select a list',
                selectedList: this.properties.lists,
                includeHidden: false,
                orderBy: PropertyFieldListPickerOrderBy.Title,
                disabled: false,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                properties: this.properties,
                context: this.context as any,
                onGetErrorMessage: null,
                deferredValidationTime: 0,
                key: 'listPickerFieldId'
              })
            ]
            }
          ]
        }
      ]
    };
  }
  
}
