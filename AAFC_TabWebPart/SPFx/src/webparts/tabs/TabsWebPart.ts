import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TabsWebPartStrings';
import Tabs from './components/Tabs';
import { ITabsProps } from './components/ITabsProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import * as $ from 'jquery';


export interface ITabsWebPartProps {
  description: string;
  sectionClass: string;
  webPartClass: string;
  tabData: any[];
}

export default class TabsWebPart extends BaseClientSideWebPart<ITabsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITabsProps> = React.createElement(
      Tabs,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        tabData: this.properties.tabData,
        webPartClass: this.properties.webPartClass,
        displayMode: this.displayMode,
        sectionClass: this.properties.sectionClass,
        jqueryDomElement: $(this.domElement)
      }
    );

    ReactDom.render(element, this.domElement);
  }

 

  private getZones(): Array<[string, string]> {
    const zones: Array<[string, string]> = [];
  
    // Get the web part ID of the current element
    const tabWebPartID = $(this.domElement)
      .closest("div." + this.properties.webPartClass)
      .attr("id");
  
    // Find the closest section (where the web part is located)
    const zoneDIV = $(this.domElement).closest("div." + this.properties.sectionClass);
  
    let count = 1;
  
    // Iterate over all web parts within the same section
    $(zoneDIV).find("." + this.properties.webPartClass).each((index, element) => {
      const thisWPID = $(element).attr("id") ?? ""; // Provide a fallback empty string in case the ID is undefined

  
      if (thisWPID !== tabWebPartID) {
        const zoneId = thisWPID;
        const zoneName: string = "Web Part " + count;
        count++;
        zones.push([zoneId, zoneName]);
      }
    });
  
    return zones;
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
                PropertyPaneTextField('sectionClass', {
                  label: strings.SectionClass,
                  description: "Class identifier for Page Section, don't touch this if you don't know what it means."
                }),
                PropertyPaneTextField('webPartClass', {
                  label: strings.WebPartClass,
                  description: "Class identifier for Web Part, don't touch this if you don't know what it means."
                }),
                PropertyFieldCollectionData("tabData", {
                  key: "tabData",
                  label: strings.TabLabels,
                  panelHeader: "Specify Labels for Tabs",
                  manageBtnLabel: "Manage Tab Labels",
                  value: this.properties.tabData,
                  fields: [
                    {
                      id: "WebPartID",
                      title: "Web Part",
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: this.getZones().map((zone:[string,string]) => {
                        return {
                          key: zone["0"],
                          text: zone["1"],
                        };
                      })

                    },
                    {
                      id: "TabLabel",
                      title: "Tab Label",
                      type: CustomCollectionFieldType.string
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}