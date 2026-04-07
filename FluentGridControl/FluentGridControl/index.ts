import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { FluentGrid } from "./FluentGrid";

declare const ReactDOM: {
    render: (element: React.ReactElement, container: Element) => void;
    unmountComponentAtNode: (container: Element) => boolean;
};

interface RecordType {
    id: string;
    product: string;
    quantity: number;
    amount: number;
    status: 'Open' | 'Active' | 'Resolved';
    category: string;
    isDirty?: boolean;
}

// ✅ Wrapper (IMPORTANT)
class GridWrapper extends React.Component<{ data: RecordType[] }> {
    render() {
        return React.createElement(FluentGrid, {
            data: this.props.data
        });
    }
}

export class FluentGridControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private container: HTMLDivElement;
    private isRendered = false;
    private lastDataHash = "";

    constructor() {
        // PCF control constructor
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        console.log("HI4 dsdasdsf sdfdf");
        const dataset = context.parameters.dataset;
        if (dataset.loading || dataset.error) return;

        const datasetChanged = context.updatedProperties.indexOf("dataset") > -1;

        if (!datasetChanged && this.isRendered) return;

        let records: RecordType[] = [];

        if (dataset.sortedRecordIds.length > 0) {
            records = dataset.sortedRecordIds.map(id => {
                const record = dataset.records[id];
                const productRef = record.getValue("productname") as ComponentFramework.LookupValue[] | null;
                const productName =
                    productRef?.[0]?.name ??
                    String(record.getValue("productname") ?? "N/A");

                const categoryRef = record.getValue("new_categoryid") as ComponentFramework.LookupValue | null;
                const categoryName = categoryRef?.name ?? "General";

                return {
                    id: id,
                    product: productName,
                    quantity: Number(record.getValue("quantity") ?? 0),
                    amount: Number(record.getValue("extendedamount") ?? 0),
                    status: 'Open' as const,
                    category: categoryName,
                    isDirty: false
                };
            });
        }

        // 🔥 IMPORTANT: Prevent re-render loop
        const newHash = JSON.stringify(records);

        if (this.lastDataHash === newHash) {
            return; // 🚫 STOP LOOP
        }

        this.lastDataHash = newHash;

        ReactDOM.render(
            React.createElement(GridWrapper, {
                data: records
            }),
            this.container
        );
    }

    public destroy(): void {
        ReactDOM.unmountComponentAtNode(this.container);
    }
}