import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxCreateElemSpecWebPartStrings';
import SpfxCreateElemSpec from './components/SpfxCreateElemSpec';
import { ISpfxCreateElemSpecProps } from './components/ISpfxCreateElemSpecProps';

import SharePointService from '../../services/SharePoint/SharePointService';

import { sp } from "@pnp/sp";
import { Environment } from "@microsoft/sp-core-library";

export interface ISpfxCreateElemSpecWebPartProps {
  description: string;
}

export default class SpfxCreateElemSpecWebPart extends BaseClientSideWebPart<ISpfxCreateElemSpecWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxCreateElemSpecProps > = React.createElement(
      SpfxCreateElemSpec,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }
  public onInit(): Promise<void> {
    return super.onInit().then(() =>{
      //let elemSpecListID = '3031e278-aab5-4dc1-aa9b-0d735b49cf29';
      let elemSpecListID = 'Element%20spec';

      //let ideaListID = 'cbcf314d-6d42-48ab-aff9-87c01bba1046';
      let ideaListID = 'Idea';
      SharePointService.setup(this.context, Environment.type, elemSpecListID, ideaListID);
      sp.setup({
        spfxContext: this.context
      });

  });}

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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
