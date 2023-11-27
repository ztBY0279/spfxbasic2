import {SPHttpClient,SPHttpClientResponse} from "@microsoft/sp-http"

function createSubsite(): void {
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
      });
  }

  export {createSubsite};