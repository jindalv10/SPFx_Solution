import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IDynamicDataCallables } from '@microsoft/sp-dynamic-data';

import * as strings from 'NominationWebPartStrings';
import Nomination from './components/Nomination';
import { NominationLibrary } from 'pd-nomination-library';
import { INominationProps } from './components/INominationProps';
import { ContextualMenuItemType } from '@microsoft/office-ui-fabric-react-bundle';
import { FormName } from './components/models/IIntakeFormDetails';

export interface INominationWebPartProps {
  description: string;
  formType:string;
}

export default class NominationWebPart extends BaseClientSideWebPart<INominationWebPartProps> {
 

  /***************************************************************************
   * Library  Component Service used to perform REST calls
   ***************************************************************************/
  private NominationQueryService: NominationLibrary;


  /***************************************************************************
   * Initializes the WebPart
  ***************************************************************************/
  protected  onInit(): Promise<void> {
    return new Promise<void>((resolve, _reject) => {
      this.NominationQueryService = new NominationLibrary(this.context);
      resolve();
    });
  }

  private _openPropertyPane(): void {
    this.context.propertyPane.open();
  }
  public render(): void {
    const element: React.ReactElement<INominationProps> = React.createElement(
      Nomination,
      {
        description: this.properties.description,
        context:this.context,
        currentUser: this.context.pageContext.user.email,
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
                
                /*
                PropertyPaneDropdown('formType', {
                  label: "Select Form",
                  options:[
                    {key: FormName.Intake, text:'Intake Form'},
                    {key: FormName.LocalAdmin, text:'Local Admin Form'},
                    {key: FormName.QC, text:'QC Form'},
                    {key: FormName.PTPAC, text:'PTPAC Form'},
                    {key: FormName.GCSLead, text:'GCS Lead Form'},
                    
                  ],
                  selectedKey: 'Drop Down 2'
                }),
                */
              ]
            }
          ]
        }
      ]
    };
  }
}
