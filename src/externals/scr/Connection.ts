import { ElementModifier, func } from '.';
import { SPHttpClient } from '@microsoft/sp-http';

class Connection {
    private context: any;

    constructor(params) {
        Object.keys(params).map(key => {
            this[key] = params[key];
        });

        this.context = this['sharepoint'].context;
    }

    public find(params) {
        return this.context.spHttpClient.get(params.link + `/_api/web/lists/getbytitle('${params.list}')/items?$select=${params.data}`, SPHttpClient.configurations.v1)
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

    public put(params){
        
    }
}

export { Connection };