import { ElementModifier, func } from '.';
import { SPHttpClient } from '@microsoft/sp-http';

class Connection {
    private context: any;
    private elementModifier;
    constructor(params) {
        Object.keys(params).map(key => {
            this[key] = params[key];
        });

        this.context = this['sharepoint'].context;
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

    private worker(params) {
        if (Worker) {
            return new Promise((resolve, reject) => {
                var working = new Worker('MainWorker.js');
                console.log(working);

                working.onmessage = event => {
                    resolve(event.data);
                };
                working.onerror = event => {
                    console.log(params);
                    reject(event);
                };

                working.postMessage(params);
            });
        }
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

    public put(params) {

    }
}

export { Connection };