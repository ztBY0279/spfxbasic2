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

// gantt chart code:- 

// import JavaScriptGantt
import { Gantt} from '@syncfusion/ej2-gantt';
 
// add Syncfusion JavaScriptstyle reference from node_modules
require('../../../node_modules/@syncfusion/ej2/fluent.css');

export interface IHelloWorldWebPartProps {
  description: string;
}


// now importing create subsite button from the external .ts file:-

//import {createSubsite} from "./helper1"

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

    <p>this is working now.</p>

    <tabel>
     <tr>
      <td>first data source</td>
     </tr>
     <tr>
      <td>second data source.</td>
     </tr>
    
    </tabel>

    <div id="Gantt-${this.instanceId}"> </div>


    <div>

    <p>changes has been made successfully.</p>

    <label>Enter Site title.</label><br>
    <input id = "title" type="text"><br>
    <label>Enter Site description.</label><br>
    <input id ="description" type="text"><br>
    <label>Enter Site URL.</label><br>
    <input id="url" type="text"><br>
    <button id="enterbutton"> create subsite </button>
    
    
    
    
    </div>


    </section>`;
    
    let data: Object[]  = [
      {
         TaskID: 1,
         TaskName: 'Project Initiation',
         StartDate: new Date('04/02/2023'),
         EndDate: new Date('04/21/2023'),
         subtasks: [
            { TaskID: 2, TaskName: 'Identify Site location', StartDate: new 
   Date('04/02/2023'), Duration: 4, Progress: 50 },
            { TaskID: 3, TaskName: 'Perform Soil test', StartDate: new Date('04/02/2023'), Duration: 4, Progress: 50  },
            { TaskID: 4, TaskName: 'Soil test approval', StartDate: new Date('04/02/2023'), Duration: 4 , Progress: 50 },
         ]
       },
      
      
   ];
    
   let gantt: Gantt = new Gantt({
      dataSource: data,
      taskFields: {
         id: 'TaskID',
         name: 'TaskName',
         startDate: 'StartDate',
         duration: 'Duration',
         dependency: 'Predecessor',
         progress: 'Progress',
         child: 'subtasks',
       },
   });
    
   gantt.appendTo('#Gantt-'+this.instanceId);
   
  


    this.postingData();
    this.subsitebtn();
   
  }

// now creatting subsite:-

private subsitebtn():void{
    
  const btn = this.domElement.querySelector("#enterbutton") as HTMLButtonElement;
  btn?.addEventListener("click",()=>{
      
    //createSubsite1();
    this.createSubsite1();
  })
}


private createSubsite1():void{
  const title1 = this.domElement.querySelector("#title") as HTMLInputElement;
  const title = title1.value;
 // const description = (document.getElementById('subsiteDescription') as HTMLInputElement).value;
 const description = (this.domElement.querySelector("#description") as HTMLInputElement).value
  const url = (this.domElement.querySelector("#url") as HTMLInputElement).value
// let temp = document.getElementById("title") as HTMLInputElement
  // Use SharePoint REST API to create a subsite
  const siteUrl: string = this.context.pageContext.web.absoluteUrl;
  const endpoint: string = `${siteUrl}/_api/web/webinfos/add`;

  const requestOptions: any = {
    headers: {
      'Accept': 'application/json;odata=verbose',
      'Content-type': 'application/json;odata=verbose',
      'odata-version': ''
    },
    body: JSON.stringify({
      'parameters': {
        '__metadata': { 'type': 'SP.WebInfoCreationInformation' },
        'Url': url,
        'Title': title,
        'Description': description,
        'Language': 1033,
        'WebTemplate': 'STS#0',
        'UseUniquePermissions': false
      }
    })
  };

  this.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, requestOptions)
    .then((response: SPHttpClientResponse) => {
      if (response.ok) {
        alert('Subsite created successfully!');
      } else {
        alert(`Error creating subsite: ${response.statusText}`);
      }
    }).catch((error)=>{
      console.log("this is error",error)
    });


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
