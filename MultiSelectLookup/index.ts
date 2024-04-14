import { getPropsWithDefaults } from "@fluentui/react";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IMultiSelectProps, MultiselectWithTags } from "./MultiSelect";
import * as React from "react";

export class MultiSelectLookup implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        const props: IMultiSelectProps = {
            utils: context.utils,
            webApi: context.webAPI,
            currentUserId: context.userSettings.userId,
            clientURL: (<any>context).page.getClientUrl(),
            addNewCallback: this.addNewCallback.bind(this),
            allowAddNew: context.parameters.allowAddNew.raw,
            items: context.parameters.dataset,
            navigateToRecord: this.navigateToRecord.bind(this),
            theme: context.fluentDesignLanguage?.tokenTheme,
            label: context.parameters.label.raw,
            thisTableName: (<any>context).page.entityTypeName,
            thisRecordId: (<any>context).page.entityId,
            relationshipName: context.parameters.relationship.raw!,
            labelLocation: context.parameters.labelLocation.raw === "0" ? "above" : "left",
            labelWidth: context.parameters.labelWidth.raw??"",
            searchMode: context.parameters.searchMode.raw == "0" ? "simple" : "advanced",
            filter: context.parameters.customFilter.raw,
            order: context.parameters.customOrder.raw,
            bestEffort: false,
            searchColumns: '',
            matchWords: 'all',
         };
        return React.createElement(
            MultiselectWithTags, props
        );
    }

    public addNewCallback(): void {
        let options: ComponentFramework.NavigationApi.EntityFormOptions = {
            entityName: this.context.parameters.dataset.getTargetEntityType(),
            useQuickCreateForm: true,
        }

        this.context.navigation.openForm(options);
    }

    public navigateToRecord(tableName: string, entityId: string): void {
        let opts: ComponentFramework.NavigationApi.EntityFormOptions = {
            entityName: tableName,
            entityId: entityId,
        };
        this.context.navigation.openForm(opts);
    }
    
    public getOutputs(): IOutputs {
        return { };
    }

    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
