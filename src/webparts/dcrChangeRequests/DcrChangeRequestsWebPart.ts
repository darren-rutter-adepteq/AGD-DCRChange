import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DcrChangeRequestsWebPartStrings';
import DcrChangeRequests from './components/DcrChangeRequests';
import { IDcrChangeRequestsProps } from './components/IDcrChangeRequestsProps';

import {sp} from "@pnp/sp";

export interface IDcrChangeRequestsWebPartProps {
  description: string;
  itemsPerPage: number;
}

export default class DcrChangeRequestsWebPart extends BaseClientSideWebPart <IDcrChangeRequestsWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IDcrChangeRequestsProps> = React.createElement(
      DcrChangeRequests,
      {
        description: this.properties.description,
        itemsPerPage: this.properties.itemsPerPage,
        siteurl: this.context.pageContext.web.absoluteUrl,
        context: this.context,
        prioritySelectedKey: "Low",
        spWebUrl: this.context.pageContext.web.absoluteUrl
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
                PropertyPaneSlider('itemsPerPage', {
                  label: "Items per page",
                  min: 5,
                  max: 20,
                  value: 10,
                  showValue: true,
                  step: 1
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
