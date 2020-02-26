import { ElementModifier, Func } from '.';
import { ISPHttpClientOptions, SPHttpClient, AadHttpClient, HttpClientResponse, MSGraphClient, IHttpClientOptions } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

class Connection {
    private context: any;
    private elementModifier;
    private func = new Func();
    constructor(params) {
        Object.keys(params).map(key => {
            this[key] = params[key];
        });

        // this.context = this['sharepoint'].context;
        this.elementModifier = new ElementModifier();
    }

    public find(params) {
        params.format = this.func.isset(params.format) ? params.format : true;

        let url = params.link + `/_api/web/lists/getbytitle('${params.list}')/items`;
        if (this.func.isset(params.data)) {
            url += `?$select=${params.data}`;
        }

        if (this.func.isset(params.filter)) {
            url += (this.func.isset(params.data)) ? '&' : '?';
            url += `$filter`;
            for (let i in params.filter) {
                url += `=${i} eq '${params.filter[i]}'`;
            }
        }

        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then(response => {
                if (response.status == 404) {
                    return 'Not Found';
                } else {
                    return response.json();
                }
            })
            .then(jsonResponse => {
                if (jsonResponse == 'Not Found') {
                    return jsonResponse;
                }
                else {
                    if (params.format) {
                        let value = [];
                        jsonResponse.value.map(row => {
                            let aRow = {};
                            for (const cell in row) {
                                if (cell.indexOf('@odata') == -1) aRow[cell.toLowerCase()] = row[cell];
                            }
                            value.push(aRow);
                        });
                        return value;
                    } else {
                        return jsonResponse.value;
                    }
                }
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
            if (this.func.isset(params.data)) request.send(params.data);
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

    public getItemEntityType(params) {
        let url = params.link + `/_api/web/lists/getbytitle('${params.list}')?$select=ListItemEntityTypeFullName`;

        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then(response => {
                return response.json();
            })
            .then(jsonResponse => {
                return jsonResponse.ListItemEntityTypeFullName;
            });
    }

    public createList(params) {
        let url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists`;

        const request: ISPHttpClientOptions = {};
        params = params || { Title: 'Sample' };
        request.body = JSON.stringify(params);

        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, request)
            .then(response => {
                if (response.status == 201) {
                    return 'Successful';
                }
                else {
                    return 'Failed';
                }
            });
    }

    public put(params) {
        let url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${params.list}')/items`;

        return this.getItemEntityType(params).then(spEntityType => {
            const request: any = {};
            params.data['@odata.type'] = spEntityType;
            request.body = JSON.stringify(params.data);

            return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, request).then(res => {
                return res.ok;
            });
        });
    }

    public update(params) {
        return this.find({ link: params.link, list: params.list, filter: params.filter, format: false }).then(stored => {
            let item = stored[0];
            let url = params.link + `/_api/web/lists/getbytitle('${params.list}')/items(${item.Id})`;

            let request: any = {};
            request.headers = {
                'X-HTTP-Method': 'MERGE',
                'IF-MATCH': (item as any)['@odata.etag']
            };

            // for (let i in params.data) {
            //     if (i.indexOf('@odata') == -1) item[i] = params.data[i];
            // }

            request.body = JSON.stringify(item);

            return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, request);
        });
    }
}

export { Connection };