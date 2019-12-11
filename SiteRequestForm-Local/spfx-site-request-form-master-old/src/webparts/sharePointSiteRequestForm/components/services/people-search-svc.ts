import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';  


class PeopleSearchService {
    private wpContext: WebPartContext;
    constructor(props) {
        this.wpContext = props;
    }

    public getSuggestions(searchVal): Promise<any> {
        return new Promise((resolve, reject) => {
            this.wpContext.spHttpClient.post(`${this.wpContext.pageContext.site.absoluteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`, 
            SPHttpClient.configurations.v1, 
            {
              credentials: 'same-origin',
                body: JSON.stringify({
                  queryParams: {
                    AllowEmailAddresses: true,
                    AllowMultipleEntities: false,
                    AllUrlZones: false,
                    MaximumEntitySuggestions: 50,
                    PrincipalSource: 15,
                    PrincipalType: 1,
                    QueryString: searchVal
                  }
                })
            }
            ).then(res => {
              res.json().then(jres => {
                resolve(jres);
              });
            });
        });
    }
}

export default PeopleSearchService;