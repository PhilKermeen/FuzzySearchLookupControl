import { IPersonaProps } from "@fluentui/react";
import * as Models from "../Models/models";

export class Mappers {


    public static MapSearchResultToPersonaProps(responses: Models.LookupResult[], primaryAtt:string): IPersonaProps[] {
        
        const result: IPersonaProps[] = [];

        responses.forEach(r => {

            result.push({
                //@ts-expect-error refdynamicsobject
                text: r.Document[primaryAtt],
                //@ts-expect-error refdynamicsobject
                id: r.Document['@search.objectid']            })
        });

        return result;
    }

 
}