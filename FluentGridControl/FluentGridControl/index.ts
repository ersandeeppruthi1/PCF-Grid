import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { FluentGrid } from "./FluentGrid";

// Declare global ReactDOM for PCF environment
declare const ReactDOM: {
    render: (element: React.ReactElement, container: Element) => void;
    unmountComponentAtNode: (container: Element) => boolean;
};

import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

interface RecordType {
    id: string;
    product: string;
    quantity: number;
    amount: number;
}

export class FluentGridControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;

    constructor() {
        console.log("FluentGridControl: Constructor called");
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        console.log("FluentGridControl: init() called");
        console.log("FluentGridControl: context:", context);
        console.log("FluentGridControl: container:", container);
        this.container = container;
        // Initialize the control
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        console.log("Hi2");

        const dataset = context.parameters.dataset;
        let records: RecordType[] = [];

        if (!dataset.loading && dataset.sortedRecordIds.length > 0) {
            records = dataset.sortedRecordIds.map(id => {
                const record = dataset.records[id];
                const productRef = record.getValue("productname") as ComponentFramework.LookupValue[] | null;
                const productName = productRef?.[0]?.name ?? String(record.getValue("productname") ?? "N/A");
                const quantity = Number(record.getValue("quantity") ?? 0);
                const amount = Number(record.getValue("extendedamount") ?? 0);

                return {
                    id: id,
                    product: productName,
                    quantity: quantity,
                    amount: amount
                };
            });
        }

        console.log("FluentGridControl: Creating React element with FluentGrid");
        console.log("FluentGridControl: Records being passed to FluentGrid:", records);
        try {
            // Render the React component to the container using React 16
            ReactDOM.render(
                React.createElement(FluentGrid, {
                    data: records,
                    context: context
                }),
                this.container
            );
        }
        catch (error) {
            console.error("FluentGridControl: Error rendering React component:", error);
        }
        console.log("FluentGridControl: React component rendered to container");
    }

    public destroy(): void {
        console.log("FluentGridControl: destroy() called");
        ReactDOM.unmountComponentAtNode(this.container);
    }
}
