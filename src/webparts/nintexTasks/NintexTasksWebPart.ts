import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'NintexTasksWebPartStrings';
import NintexTasks from './components/NintexTasks';
import { INintexTasksProps } from './components/INintexTasksProps';
import { TaskStatus } from './models/INintexTask';

export interface INintexTasksWebPartProps {
  tenantUrl: string;
  tokenListUrl: string;
  tokenTitleFilter: string;
  tokenColumnName: string;
  dashboardUrl: string;
  defaultStatusFilter: TaskStatus;
  headerText: string;
  headerBgColor: string;
  headerTextColor: string;
  useDefaultColors: boolean;
}

export default class NintexTasksWebPart extends BaseClientSideWebPart<INintexTasksWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<INintexTasksProps> = React.createElement(
      NintexTasks,
      {
        tenantUrl: this.properties.tenantUrl || '',
        tokenListUrl: this.properties.tokenListUrl || '',
        tokenTitleFilter: this.properties.tokenTitleFilter || '',
        tokenColumnName: this.properties.tokenColumnName || 'Token',
        dashboardUrl: this.properties.dashboardUrl || '',
        defaultStatusFilter: this.properties.defaultStatusFilter || 'active',
        httpClient: this.context.httpClient,
        spHttpClient: this.context.spHttpClient,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        headerText: this.properties.headerText,
        headerBgColor: this.properties.headerBgColor,
        headerTextColor: this.properties.headerTextColor,
        useDefaultColors: this.properties.useDefaultColors
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit();
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
    const statusOptions: IPropertyPaneDropdownOption[] = [
      { key: 'active', text: 'Active' },
      { key: 'complete', text: 'Complete' },
      { key: 'expired', text: 'Expired' },
      { key: 'overridden', text: 'Overridden' },
      { key: 'terminated', text: 'Terminated' },
      { key: 'all', text: 'All' }
    ];

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
                PropertyPaneTextField('tenantUrl', {
                  label: 'Nintex API Base URL',
                  placeholder: 'https://us.nintex.io',
                  description: 'The regional API endpoint for Nintex (e.g., https://us.nintex.io)'
                }),
                PropertyPaneTextField('dashboardUrl', {
                  label: 'Nintex Dashboard base url',
                  placeholder: 'https://mydev26-507348.workflowcloud.com',
                  description: 'The NWC Dashboard base URL for task deep links'
                }),
                PropertyPaneTextField('tokenListUrl', {
                  label: 'Token List URL',
                  placeholder: "https://tenant.sharepoint.com/sites/mysite/_api/web/lists/getbytitle('NintexTokens')/items",
                  description: 'Full SP REST API URL to the list items endpoint (must include /_api/... and /items)'
                }),
                PropertyPaneTextField('tokenTitleFilter', {
                  label: 'Token List Title Filter',
                  placeholder: "NWC_TOKEN",
                  description: 'Value to filter the Title column by in the token list'
                }),
                PropertyPaneTextField('tokenColumnName', {
                  label: 'Token Column Name',
                  placeholder: "Token",
                  description: 'Internal name of the column containing the token value (default: Token)'
                }),
                PropertyPaneDropdown('defaultStatusFilter', {
                  label: 'Default Status Filter',
                  options: statusOptions,
                  selectedKey: 'active'
                })
              ]
            },
            {
              groupName: 'Header Configuration',
              groupFields: [
                PropertyPaneTextField('headerText', {
                  label: 'Header Text',
                  placeholder: 'Assigned Tasks',
                  description: 'Custom title for the web part header (default: Assigned Tasks)'
                }),
                PropertyPaneToggle('useDefaultColors', {
                  label: 'Use Default Colors',
                  checked: this.properties.useDefaultColors !== false
                }),
                PropertyPaneTextField('headerBgColor', {
                  label: 'Header Background Color',
                  placeholder: '#8c39df',
                  description: 'Hex color code (e.g., #8c39df) for the header background',
                  disabled: this.properties.useDefaultColors !== false
                }),
                PropertyPaneTextField('headerTextColor', {
                  label: 'Header Text Color',
                  placeholder: '#ffffff',
                  description: 'Hex color code (e.g., #ffffff) for the header text',
                  disabled: this.properties.useDefaultColors !== false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
