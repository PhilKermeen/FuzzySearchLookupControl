import * as Models from "../Models/models";

export class DynamicsRepository {

    public static async tryGetEntityPrimaryName(entityName: string)  {
        try {
            const response = await DynamicsRepository.getEntityPrimaryName(entityName);
            console.log(response);
            return response;
        } catch (error) {
            console.error(error);
        }
    }

    public static getEntityPrimaryName(entityName: string): Promise<string>
    {
        return new Promise((resolve, reject) => {
            const req = new XMLHttpRequest();
            req.open("GET", "/api/data/v9.2/EntityDefinitions(LogicalName='" + entityName + "')?$select=PrimaryNameAttribute", true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Prefer", "odata.include-annotations=*");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200 || this.status === 204) {
                        const result = JSON.parse(this.response);
                        console.log(result);
                        resolve(result.PrimaryNameAttribute);

                    } else {
                        console.log(this.responseText);
                        reject(new Error('problem calling get primary name ' + this.responseText))
                    }
                }
            };
       
            req.send();
        });
    }

    public static async tryGetSearch(searchTerm: string, entityName: string)  {
        try {
            const response = await DynamicsRepository.getSearchVals(searchTerm, entityName);
            console.log(response);
            return response;
        } catch (error) {
            console.error(error);
        }
    }

    public static getSearchVals(searchTerm: string, entityName:string): Promise<Models.LookupResult[]>
    {
        return new Promise((resolve, reject) => {
            const req = new XMLHttpRequest();
            req.open("POST", "/api/data/v9.2/searchsuggest", true);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Accept", "application/json");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200 || this.status === 204) {
                        const result = JSON.parse(this.response);
                        console.log(result);
                        resolve(JSON.parse(result.response).Value);

                    } else {
                        console.log(this.responseText);
                        reject(new Error('problem calling search ' + this.responseText))
                    }
                }
            };
        
            req.send(JSON.stringify({ "fuzzy": true, "entities": "[{'name': '" + entityName + "'}]",  "search": searchTerm }));
        });
    }

  


}

