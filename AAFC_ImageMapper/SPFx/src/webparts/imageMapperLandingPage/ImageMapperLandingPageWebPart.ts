import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ImageMapperLandingPageWebPartStrings';
import ImageMapperLandingPage from './components/ImageMapperLandingPage';
import { IImageMapperLandingPageProps } from './components/IImageMapperLandingPageProps';
import { IMapArea } from '../models/IMapArea';

export interface IImageMapperLandingPageWebPartProps {
  description: string;
  imageUrl: string;
  imageHeight: string;
  imageWidth: string;
  imageHorizontalPosition: string;
  imageVerticalPosition: string;
  scale: number;
  items: IMapArea[];
}

export default class ImageMapperLandingPageWebPart extends BaseClientSideWebPart<IImageMapperLandingPageWebPartProps> {


  public render(): void {
    const element: React.ReactElement<IImageMapperLandingPageProps> = React.createElement(
      ImageMapperLandingPage,
      {
        description: this.properties.description,
        imageUrl: this.properties.imageUrl,
        imageHeight: this.properties.imageHeight,
        imageWidth: this.properties.imageWidth,
        imageHorizontalPosition: this.properties.imageHorizontalPosition,
        imageVerticalPosition: this.properties.imageVerticalPosition,
        scale: this.properties.scale,
        items: this.properties.items,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onAddButtonClick(value: any): void {
    this.properties.items.push({});
  }

  protected onDeleteButtonClick(value: any): void {
    this.properties.items.splice(value, 1);
  }

  private createNewGroup(iMapArea: IMapArea, index: number): any {
    if (iMapArea.imapType === "Path") {
      return {
        groupName: `Mapped Area ${index + 1}`,
        groupFields: [
          PropertyPaneDropdown(`items[${index}].imapType`, {
            label: "Map Area Type",
            selectedKey: iMapArea.imapType,
            options: [
              { key: "Rectangle", text: "Rectangle" },
              { key: "Path", text: "Path" },
            ],
          }),
          PropertyPaneTextField(`items[${index}].d`, {
            label: "D",
            value: iMapArea.d,
          }),
          PropertyPaneTextField(`items[${index}].url`, {
            label: "Url",
            value: iMapArea.url,
          }),
          PropertyPaneCheckbox(`items[${index}].openInNewWindow`, {
            checked: iMapArea.openInNewWindow,
            text: "Open in new window",
          }),
          PropertyPaneButton("deleteButton", {
            text: "Delete",
            buttonType: PropertyPaneButtonType.Command,
            icon: "RecycleBin",
            onClick: this.onDeleteButtonClick.bind(this, index),
          }),
          PropertyPaneButton("addButton", {
            text: "Add",
            buttonType: PropertyPaneButtonType.Command,
            icon: "CirclePlus",
            onClick: this.onAddButtonClick.bind(this),
          }),
        ],
      };
    } else {
      return {
        groupName: `Mapped Area ${index + 1}`,
        groupFields: [
          PropertyPaneDropdown(`items[${index}].imapType`, {
            label: "Map Area Type",
            selectedKey: iMapArea.imapType,
            options: [
              { key: "Rectangle", text: "Rectangle" },
              { key: "Path", text: "Path" },
            ],
          }),
          PropertyPaneTextField(`items[${index}].x`, {
            label: "X",
            value: iMapArea.x,
          }),
          PropertyPaneTextField(`items[${index}].y`, {
            label: "Y",
            value: iMapArea.y,
          }),
          PropertyPaneTextField(`items[${index}].width`, {
            label: "Width",
            value: iMapArea.width,
          }),
          PropertyPaneTextField(`items[${index}].height`, {
            label: "Height",
            value: iMapArea.height,
          }),
          PropertyPaneTextField(`items[${index}].url`, {
            label: "Url",
            value: iMapArea.url,
          }),
          PropertyPaneCheckbox(`items[${index}].openInNewWindow`, {
            checked: iMapArea.openInNewWindow,
            text: "Open in new window",
          }),
          PropertyPaneButton("deleteButton", {
            text: "Delete",
            buttonType: PropertyPaneButtonType.Command,
            icon: "RecycleBin",
            onClick: this.onDeleteButtonClick.bind(this, index),
          }),
          PropertyPaneButton("addButton", {
            text: "Add",
            buttonType: PropertyPaneButtonType.Command,
            icon: "CirclePlus",
            onClick: this.onAddButtonClick.bind(this),
          }),
        ],
      };
    }
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      console.log(message);  // Example action with message
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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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
    const pages = [];

    pages.push({
      header: {
        description: "Image Area Settings",
      },
      groups: [
        {
          groupName: "Image",
          groupFields: [
            PropertyPaneTextField("imageUrl", {
              label: "Image Url",
            }),
            PropertyPaneTextField("imageHeight", {
              label: "Image Height",
            }),
            PropertyPaneTextField("imageWidth", {
              label: "Image Width",
            }),
            PropertyPaneDropdown("imageHorizontalPosition", {
              label: "Image Horizontal Position",
              options: [
                { key: "left", text: "Left" },
                { key: "center", text: "Center" },
                { key: "right", text: "Right" },
              ],
            }),
            PropertyPaneDropdown("imageVerticalPosition", {
              label: "Image Vertical Position",
              options: [
                { key: "start", text: "Top" },
                { key: "center", text: "Center" },
                { key: "end", text: "Bottom" },
              ],
            }),
            PropertyPaneSlider("scale", {
              label: "Scale",
              min: 0,
              max: 100,
            }),
          ],
        },
        {
          groupName: "",
          groupFields: [
            PropertyPaneButton("addButton", {
              text: "Add",
              buttonType: PropertyPaneButtonType.Command,
              icon: "CirclePlus",
              onClick: this.onAddButtonClick.bind(this),
            }),
          ],
        },
      ],
    });

    console.log(this.properties);
    this.properties.items.forEach((item, index) => {
      pages.push({
        header: {
          description: `Map Area ${index + 1}`,
        },
        groups: [this.createNewGroup(item, index)],
      });
    });

    return {
      pages: pages,
    };
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (propertyPath === "imageUrl" && newValue) {
      const image = new Image();
      image.src = newValue;

      image.onload = () => {
        this.properties.imageHeight = image.height.toString();
        this.properties.imageWidth = image.width.toString();
      };
    }
  }
}