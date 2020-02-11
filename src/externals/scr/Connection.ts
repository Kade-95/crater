import { ElementModifier, func } from '.';
import { SPHttpClient, AadHttpClient, HttpClientResponse, MSGraphClient, IHttpClientOptions } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

class Connection {
    private context: any;
    private elementModifier;
    constructor(params) {
        Object.keys(params).map(key => {
            this[key] = params[key];
        });

        // this.context = this['sharepoint'].context;
        this.elementModifier = new ElementModifier();
    }

    public find(params) {
        let url = params.link + `/_api/web/lists/getbytitle('${params.list}')/items`;
        if (func.isset(params.data)) {
            url += `?$select=${params.data}`;
        }

        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then(response => {
                return response.json();
            })
            .then(jsonResponse => {
                let value = [];
                console.log(jsonResponse);

                jsonResponse.value.map(row => {
                    let aRow = {};
                    for (const cell in row) {
                        if (cell.indexOf('@odata') == -1) aRow[cell.toLowerCase()] = row[cell];
                    }
                    value.push(aRow);
                });
                return value;
            });
    }

    public ajax(params) {
        params.async = params.async || true;
        return new Promise((resolve, reject) => {
            var result;
            var request = new XMLHttpRequest();
            request.onreadystatechange = function (e) {
                if (this.readyState == 4 && this.status == 200) {
                    resolve(request.responseText);
                }
            };

            request.open(params.method, params.url, params.async);
            if (func.isset(params.data)) request.send(params.data);
            else request.send();
        });
    }

    public getSite() {
        let site = '';

        if (location.pathname == '') {
            site = location.origin;
        }
        else if (location.pathname.split('/').indexOf('SitePages') == 1) {
            site = location.origin;
        }
        else if (location.pathname.split('/').indexOf('sites') == 1) {
            site = location.origin + '/' + location.pathname.split('/')[1] + '/' + location.pathname.split('/')[2];
        }

        return site;
    }

    public getWithAad(endPoint, url, options?) {
        return new Promise((resolve, reject) => {
            this.context.aadHttpClientFactory.getClient(url)
                .then((aadClient: AadHttpClient) => {
                    aadClient.get(endPoint, AadHttpClient.configurations.v1)
                        .then((rawResponse: HttpClientResponse) => {
                            return rawResponse.json();
                        })
                        .then((jsonResponse: any) => {
                            resolve(jsonResponse);
                        });
                });
        });
    }

    public updateWithAad(endPoint, url, options: IHttpClientOptions) {
        return new Promise((resolve, reject) => {
            this.context.aadHttpClientFactory.getClient(url)
                .then((aadClient: AadHttpClient) => {
                    aadClient.post(endPoint, AadHttpClient.configurations.v1, options)
                        .then((rawResponse: HttpClientResponse) => {
                            return rawResponse.json();
                        })
                        .then((jsonResponse: any) => {
                            resolve(jsonResponse);
                        });
                });
        });
    }

    public getWithGraph() {
        return new Promise((resolve, reject) => {
            this.context.msGraphClientFactory.getClient()
                .then((client: MSGraphClient): void => {
                    resolve(client);
                });
        });
    }

    public put(params) {

    }
}

export { Connection };