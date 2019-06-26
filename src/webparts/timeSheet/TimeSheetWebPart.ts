import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneCheckbox } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'TimeSheetWebPartStrings';
import TimeSheet from './components/TimeSheet';
import { ITimeSheetProps } from './components/TimeSheet';

/*
NOTE: Hacky fix, but in order for the dx themes to work, all @font-face blocks need to have the woff2 reference rmeoved
*/
// import 'jquery/dist/jquery.min.js';
import 'devextreme/dist/css/dx.common.css';
import 'devextreme/dist/css/dx.light.css';
import 'bootstrap/dist/css/bootstrap.min.css';
// import 'devextreme/dist/js/dx.all.js';
// import 'bootstrap/dist/js/bootstrap.min.js';
/*
  Have to remove woff2 references from fontawesome too
*/
import 'font-awesome/css/font-awesome.css';

export interface ITimeSheetWebPartProps {
  admin: boolean;
  devMode: boolean;
}

export default class TimeSheetWebPart extends BaseClientSideWebPart<ITimeSheetWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITimeSheetProps > = React.createElement(
      TimeSheet,
      {
        admin: this.properties.admin === true,
        devMode: this.properties.devMode === true
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
                PropertyPaneCheckbox('admin', {
                  text: "Admin View",
                  checked: this.properties.admin === true 
                }),
                PropertyPaneCheckbox('devMode', {
                  text: "Developer Mode",
                  checked: this.properties.devMode === true 
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
