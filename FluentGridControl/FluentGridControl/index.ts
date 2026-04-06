import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { FluentGrid } from "./FluentGrid";

import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

interface RecordType {
    id: string;
    product: string;
    quantity: number;
    amount: number;
}

export class FluentGridControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  
    constructor() {
        // Empty
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        // Initialize the control
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        const dataset = context.parameters.dataset;
        let records: RecordType[] = [];

        if (!dataset.loading && dataset.sortedRecordIds.length > 0) {
            records = dataset.sortedRecordIds.map(id => {
                const record = dataset.records[id];
                const productRef = record.getValue("productname") as ComponentFramework.LookupValue[] | null;
                return {
                    id: id,
                    product: productRef?.[0]?.name ?? String(record.getValue("productname") ?? "N/A"),
                    quantity: Number(record.getValue("quantity") ?? 0),
                    amount: Number(record.getValue("extendedamount") ?? 0)
                };
            });
        }

        // Return the element directly instead of calling ReactDOM.render
        return React.createElement(FluentGrid, {
            data: records,
            context: context
        });
    }

    public destroy(): void {
        // Remove ReactDOM.unmountComponentAtNode(this.container)
    }
}
