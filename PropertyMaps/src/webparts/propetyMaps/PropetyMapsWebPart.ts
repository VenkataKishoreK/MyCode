import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PropetyMapsWebPartStrings';
import PropetyMaps from './components/PropetyMaps';
import { IPropetyMapsProps } from './components/IPropetyMapsProps';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

export interface IPropetyMapsWebPartProps {
 description: string;
 lists: string|string[];
 multiSelect: string[];
}


export default class PropetyMapsWebPart extends BaseClientSideWebPart <IPropetyMapsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPropetyMapsProps> = React.createElement(
      PropetyMaps,
      {
        description: this.properties.description,
        lists: this.properties.lists,
        multiSelect: this.properties.multiSelect        
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
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldMultiSelect('multiSelect', {
                  key: 'multiSelect',
                  label: "Multi select field",
                  options: [
                    {
                      key: "EN",
                      text: "EN"
                    },
                    {
                      key: "FR",
                      text: "FR"
                    },
                    {
                      key: "NL",
                      text: "NL"
                    }
                  ],
                  selectedKeys: this.properties.multiSelect
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
