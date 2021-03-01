import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './BasicDetailsWebPart.module.scss';
import * as strings from 'BasicDetailsWebPartStrings';

export interface IBasicDetailsWebPartProps {
  description: string;
}

export default class BasicDetailsWebPart extends BaseClientSideWebPart <IBasicDetailsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.basicDetails }">
    <div class="${ styles.container }">
      <div class="${ styles.row }">
        <div class="${ styles.column }">

        <p class="${ styles.description } displayName">Welcome to Sharepoint!</p>


     
          </div>
          </div>
          </div>
          </div>`;
          this.getuserDetails();
  }
 async getuserDetails()
{

  await this.context.msGraphClientFactory.getClient().then(client => {
    client.api("/me").select("*").get((error, response) => {
      var getResponse = response;
      console.log("My Details: ");
      console.log(response)
      client.api("/me/manager").get((error, response) => {
        var getmanagerResponse = response;
        console.log("My manager:")
        console.log(getmanagerResponse)
  
  
    });
    client.api("/me/manager").select("department").get((error, response) => {
      var getDepartment = response;
      console.log("My department:")
      console.log(getDepartment)


  });

  })
});
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
