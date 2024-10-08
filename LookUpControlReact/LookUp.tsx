import * as React from 'react';
import { IPersonaProps, Label, NormalPeoplePicker } from '@fluentui/react';
import { LookUpControlReact } from '.';

import { DynamicsRepository } from './Repositories/dynamics-repository';
import { Mappers } from './Models/Mappers';

export interface ILookUpProps {
  name: string;
  startEntityId: string;
  startEntityName:string;
  onEntityChange: (selectedOption?: ComponentFramework.LookupValue[] | undefined) => void;
}

export class LookUp extends React.Component<ILookUpProps> {

  public primaryAttr:string;

  public onFilterChanged =  async (filter: string, selectedItems?: IPersonaProps[] | undefined): Promise<IPersonaProps[]> => {
    return this.getSearchResultsFromDynamics(filter);
  }

  private getSearchResultsFromDynamics = async (filterText: string): Promise<IPersonaProps[]> => {
    try {

      if(this.primaryAttr==undefined)
      {
        DynamicsRepository.tryGetEntityPrimaryName(this.props.name).then(r=>{this.primaryAttr = r!});
      }

      if(filterText.length > 2)
      {
        const searchResults = await DynamicsRepository.tryGetSearch(filterText.toLowerCase(), this.props.name);

        if (searchResults) {
            return Mappers.MapSearchResultToPersonaProps(searchResults, this.primaryAttr);
        }
      }
      return [];
    } catch (error) {
        alert(`An error has occurred when trying to contact Search - Please contact your administrator`);
        console.log(`An error has occurred when trying to search: ${error}`);
        throw error;
    }
};

private loadCurrentRecord = ():IPersonaProps[] =>{

  const currentRecord: IPersonaProps[] = [];
  if(this.props.startEntityId != undefined){
    currentRecord.push({text: this.props.startEntityName, id: this.props.startEntityId});
  }
  return currentRecord;
}

private  onChangeSelectedEntity = (e?: IPersonaProps[]) => {
  const lookupvalues: ComponentFramework.LookupValue[] = [];
  if(e != null && e.length != 0 ){
    const retVal : ComponentFramework.LookupValue = {id: e[0].id! , name: e[0].text, entityType: this.props.name}
     lookupvalues.push( retVal ); 
  }
  this.props.onEntityChange(lookupvalues);
};

  public render(): React.ReactNode {
    return (
      <NormalPeoplePicker
         onResolveSuggestions={this.onFilterChanged}
         defaultSelectedItems={this.loadCurrentRecord()}
         onChange={this.onChangeSelectedEntity}
         itemLimit={1}
      />
    )
  }
}
