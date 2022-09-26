import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudWrbpartWebPart.module.scss';
import * as strings from 'CrudWrbpartWebPartStrings';

export interface ICrudWrbpartWebPartProps {
  description: string;
}

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

interface IRegistrationDetails {
  Title: string;
  Address: string;
  Mobile: number;
  EmailID: string;
  }

   

export default class CrudWrbpartWebPart extends BaseClientSideWebPart<ICrudWrbpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private Listname: string = "Employee";
  private listItemId: number = 0;

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }


  public render(): void {
    this.domElement.innerHTML =`<div>
    <table>
    <tr>
    <td>Full Name</td>
    <td><input type="text" id="idFullName" name="fullName" placeholder="Full Name.."></td>
    </tr>
    <tr>
    <td>Address</td>
    <td><input type="text" id="idAddress" name="address" placeholder="Address.."></td>
    </tr>
    <tr>
    <td>Mobile Number</td>
    <td><input type="text" id="idPhoneNumber" name="mobile" placeholder="Mobile Number.."></td>
    </tr>
    <tr>
    <td>Email ID</td>
    <td><input type="text" id="idEmailId" name="emailid" placeholder="Email ID.."></td>
    </tr>
    </table>
    <table>
    <tr>
    <td><button class="${styles.button} find-Button" >Find</button></td>
    <td><button class="${styles.button} create-Button">Create</button></td>
    <td><button class="${styles.button} update-Button">Update</button></td>
    <td><button class="${styles.button} delete-Button">Delete</button></td>
    <td><button class="${styles.button} clear-Button">Clear</button></td>
    </tr>
    </table>
    <div id="tblRegistrationDetails"></div>
    </div>
    `;
    this.setButtonsEventHandlers();
    this.getListData();  
  }


  private setButtonsEventHandlers(): void {
    const webPart: CrudWrbpartWebPart = this;
    this.domElement.querySelector('button.find-Button').addEventListener('click', () => { webPart.find(); });
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.create(); });
    this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.update(); });
    this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart.delete(); });
    this.domElement.querySelector('button.clear-Button').addEventListener('click', () => { webPart.clear(); });
  }

  private find(): void {
    
    let emailId = prompt("Enter the Email ID");
    var siteUrl = this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items?$select=*&$filter=EmailID eq '${emailId}'`, SPHttpClient.configurations.v1)
    .then(response => {
    return response.json()
    .then((item: any): void => {
      console.log(siteUrl);
      (document.getElementById('idFullName') as HTMLInputElement).value =item.value[0].Title;
    //document.getElementById('idFullName')["value"] = item.value[0].Title;
    (document.getElementById('idAddress') as HTMLInputElement).value  = item.value[0].Address;
    //document.getElementById('idAddress')["value"] = item.value[0].Address;
    (document.getElementById('idEmailId') as HTMLInputElement).value = item.value[0].EmailID;
      //document.getElementById('idEmailId')["value"] = item.value[0].EmailID;

    (document.getElementById('idPhoneNumber') as HTMLInputElement).value = item.value[0].Mobile;
   // document.getElementById('idPhoneNumber')["value"] = item.value[0].Mobile;
    this.listItemId = item.value[0].Id;
    });
    });
    }
         // This Function All List Item 
    private getListData() {
      
      let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
      html += '<th>Full Name</th><th>Address</th><th>Email ID</th> <th>Phone Number</th>  ';
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items`, SPHttpClient.configurations.v1)
      .then(response => {
      return response.json()
      .then((items: any): void => {
      console.log('items.value: ', items.value);
      const listItems: IRegistrationDetails[] = items.value;
      console.log('list items: ', listItems);
       
      listItems.forEach((item: IRegistrationDetails) => {
      html += `
      <tr>
      <td>${item.Title}</td>
      <td>${item.Address}</td>
      <td>${item.EmailID}</td>
      <td>${item.Mobile}</td>
      </tr>
      `;
      });
      html += '</table>';
      const listContainer: Element = this.domElement.querySelector('#tblRegistrationDetails');
      listContainer.innerHTML = html;
      });
      });
    }
  

    private Create(): void {
     // this Method Created By Gulam Khan 
      //   Declare All Variable Globally   *//
      let  Fullname         = document.getElementById("idFullName");
      let Address           = document.getElementById("idAddress");
      let Email             = document.getElementById("idEmailId");
      let Phone_number       = document.getElementById("idPhoneNumber");
     
      const body: string = JSON.stringify({
      'Title': (Fullname as HTMLInputElement).value,
      'Address': (Address as HTMLInputElement).value,
      'EmailID': (Email as HTMLInputElement).value,
      'Mobile': (Phone_number as HTMLInputElement).value
      });
       
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items`,
      SPHttpClient.configurations.v1,
      {
      headers: {
      'Accept': 'application/json;odata=nometadata',
      'X-HTTP-Method': 'POST'
      },
      body: body
      }).then((response: SPHttpClientResponse): void => {
      this.getListData();
      this.clear();
      alert('Item has been successfully Saved ');
      }, (error: any): void => {
      alert(`${error}`);
      });
      }


      private update(): void {
       
       
     
        const body: string = JSON.stringify({
          'Title': (document.getElementById('idFullName') as HTMLInputElement).value,
          'Address': (document.getElementById('idAddress') as HTMLInputElement).value,
          'EmailID': (document.getElementById('idEmailId') as HTMLInputElement).value,
          'Mobile': (document.getElementById('idPhoneNumber') as HTMLInputElement).value
        });
         
        var current_url = this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${this.listItemId})`,
        
        SPHttpClient.configurations.v1,
        {
        headers: {
        'Accept': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'PATCH'
        },
        body: body
        }).then((response: SPHttpClientResponse): void => {
        this.getListData();
        this.clear();
        console.log(current_url);
        alert(`Item successfully updated`);
        }, (error): void => {
        alert(`${error}`);
        });
        }


        private Update(): void {
          
          if (!window.confirm('Are you sure you want to delete the latest item?')) {
          return;
          }
           
          this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${this.listItemId})`,
          SPHttpClient.configurations.v1,
          {
          headers: {
          'Accept': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
          }
          }).then((response: SPHttpClientResponse): void => {
          alert(`Item successfully Deleted`);
          this.getListData();
          this.clear();
          }, (error: any): void => {
          alert(`${error}`);
          });
          }


          private clear(): void {
             (document.getElementById('idFullName') as HTMLInputElement).value  ="";
            (document.getElementById('idAddress') as HTMLInputElement).value    ="";
            (document.getElementById('idEmailId') as HTMLInputElement).value    ="";
            (document.getElementById('idPhoneNumber') as HTMLInputElement).value  ="";
            
        }


  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
