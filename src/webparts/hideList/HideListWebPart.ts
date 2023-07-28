import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import ListeTemplateType from './assets/ListeTemplateType.json';
import LisetBaseType from  './assets/ListeBaseType.json';

import * as strings from 'HideListWebPartStrings';
import HideList from './components/HideList';
import { IHideListProps } from './components/IHideListProps';
import { getSP } from './pnpjsConfig';

import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';


export interface IHideListWebPartProps {
  ListeTemplate: string[],
  BaseType: string[];
}

export default class HideListWebPart extends BaseClientSideWebPart<IHideListWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected async onInit(): Promise<void> {
    this._getEnvironmentMessage().then(res => this._environmentMessage = res);

    // Initialize our _sp object that we can then use in other packages without having to pass around the context.
    // Check out pnpjsConfig.ts for an example of a project setup file.
    getSP(this.context);
    this.properties.BaseType = ["0"];  
    this.properties.ListeTemplate = ["100"];  
  }

  public render(): void {
    const element: React.ReactElement<IHideListProps> = React.createElement(
      HideList,
      {
        BaseType: this.properties.BaseType,
        ListeTemplate: this.properties.ListeTemplate,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
      }
    );

    ReactDom.render(element, this.domElement);
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
            description: "Page 1"
          },
          groups: [
            {
              groupName: "Tri des listes",
              groupFields: [
                
                PropertyFieldMultiSelect('BaseType',{
                  label: 'BaseType label',
                  key:'BaseType',         
                  options: LisetBaseType,
                  selectedKeys: this.properties.BaseType
                }),
                PropertyFieldMultiSelect('ListeTemplate', {
                  label: "ListeTemplate",
                  key:'ListeTemplate',
                  options:ListeTemplateType,
                  selectedKeys: this.properties.ListeTemplate
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
