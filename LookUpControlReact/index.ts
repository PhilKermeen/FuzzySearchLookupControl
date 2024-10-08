import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ILookUpProps, LookUp } from "./LookUp";
import * as React from "react";

export class LookUpControlReact implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;

    private entityName: string;
    private startEntityId: string;
    private startEntityName: string;
    private selectedValue: ComponentFramework.LookupValue[] | undefined;
    private lookupProps: ILookUpProps

    /**
     * Empty constructor.
     */
    constructor() { }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.entityName = context.parameters.lookupProperty.getTargetEntityType();
        if(context.parameters.lookupProperty.raw.length > 0){
            this.startEntityId = context.parameters.lookupProperty.raw[0].id;
            this.startEntityName = context.parameters.lookupProperty.raw[0].name!;
        }
        this.notifyOutputChanged = notifyOutputChanged;
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {

        this.lookupProps = { name: this.entityName, startEntityId: this.startEntityId, startEntityName: this.startEntityName, onEntityChange: this.onChange };
        return React.createElement(
            LookUp, this.lookupProps
        );
    }

    onChange = (newValue: ComponentFramework.LookupValue[] | undefined): void => {
        this.selectedValue = newValue
        this.notifyOutputChanged()
      };

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {
            lookupProperty: this.selectedValue
         };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
