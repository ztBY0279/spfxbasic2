import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { escape } from '@microsoft/sp-lodash-subset';

//import { Gantt, Sort } from ‘@syncfusion/ej2-gantt’;

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import { SPHttpClient, SPHttpClientResponse,ISPHttpClientOptions } from '@microsoft/sp-http';
//import { response } from 'express';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

   private _isDarkTheme: boolean = false;
   private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        
    <label>Enter List name</label>
    <input id = "input1" type="text" name = "listname"><br>
    <label>Enter description</label>
    <input id = "input2" type = "text" name = "description">
    <button id = "submit">create List</button>


    </section>`;
    this.postingData();
   
  }

  private postingData():void{
    const submitid = this.domElement.querySelector("#submit") as HTMLButtonElement;

    submitid.addEventListener("click",()=>{
      this.submit();
    })
  }

  private submit():void{
     const input1 = this.domElement.querySelector("#input1") as HTMLInputElement;

     const input2 = this.domElement.querySelector("#input2") as HTMLInputElement;

    
     

     const  listname = input1.value
     console.log(listname);
     const description = input2.value;
     console.log(input2.value);

     this.createSharepointList(listname,description);
     
    
  }

  private createSharepointList(listname:string,description:string):void{

  console.log(listname);
    const url = this.context.pageContext.web.absoluteUrl+ " /_api/web/lists/GetByTitle('" + listname + "')";

   console.log("the complete url is :",url);

    this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
    .then((response:SPHttpClientResponse)=>{
         
      if(response.status === 200){
        alert("this list is already exist:");

        return ;
      }

      if(response.status === 404){

        const endpointUrl:string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";

        const listMetadata = {
        
          "BaseTemplate": 100,  // 100 for custom list
          "Title": listname,
          "Description": description,
          "AllowContentTypes": true,
          "ContentTypesEnabled": true
        };
    
        const config:ISPHttpClientOptions = {
          "body": JSON.stringify(listMetadata)
        }
    
        this.context.spHttpClient.post(endpointUrl, SPHttpClient.configurations.v1, config)
        .then((response1: SPHttpClientResponse): void => {
          if (response1.status === 201) {
           
            alert("a new list has been created :");
          } else {
    
            alert("list is not created:");
            
          }
        }).catch((error)=>{
          console.log("this is error:",error);
        })


      }

    }).catch((error)=>{
      console.log("this is errror",error);
    })

   
   
   
    
  }


  



  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
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
