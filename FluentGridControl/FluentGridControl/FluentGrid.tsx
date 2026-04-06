import * as React from "react";
import {
    DetailsList,
    IColumn,
    SelectionMode,
    Stack,
    TextField,
    Spinner,
    DefaultButton,
    PrimaryButton
} from "@fluentui/react";
import axios from "axios";

interface RecordType {
    id: string;
    product: string;
    quantity: number;
    amount: number;
}

interface GroupHeaderType {
    isGroup: true;
    key: string;
    name: string;
    count: number;
    totalQuantity: number;
    totalAmount: number;
}

type GridItem = RecordType | GroupHeaderType;

interface FluentGridProps {
    data?: RecordType[];
    context?: ComponentFramework.Context<unknown>;
}

export const FluentGrid: React.FC<FluentGridProps> = ({ data: initialData, context }) => {

    // ---------------- STATE ----------------
    const [data, setData] = React.useState<RecordType[]>(initialData || []);
    const [loading, setLoading] = React.useState(false);

    const [request, setRequest] = React.useState({
        page: 1,
        pageSize: 25,
        filters: {} as Record<string, unknown>
    });

    const [editing, setEditing] = React.useState<Record<string, Partial<RecordType>>>({});

    // ---------------- FETCH ----------------
    const fetchData = React.useCallback(async () => {
        setLoading(true);
        try {
            const res = await axios.post("/api/grid", request);
            setData(res.data.items);
        } finally {
            setLoading(false);
        }
    }, [request]);

    React.useEffect(() => {
        const t = setTimeout(fetchData, 400);
        return () => clearTimeout(t);
    }, [fetchData]);

    // ---------------- FILTER ----------------
    const setFilter = (field: string, value: string) => {
        setRequest(prev => ({
            ...prev,
            page: 1,
            filters: {
                ...prev.filters,
                [field]: value
            }
        }));
    };

    // ---------------- EDIT ----------------
    const updateField = (id: string, field: keyof RecordType, value: string | number | undefined) => {
        if (value !== undefined) {
            setEditing(prev => ({
                ...prev,
                [id]: {
                    ...prev[id],
                    [field]: value
                }
            }));
        }
    };

    // ---------------- SAVE ----------------
    const saveRow = async (id: string) => {
        const changes = editing[id];
        if (!changes) return;

        await axios.put(`/api/grid/${id}`, changes);

        setData(prev =>
            prev.map(x => (x.id === id ? { ...x, ...changes } : x))
        );

        setEditing(prev => {
            const copy = { ...prev };
            delete copy[id];
            return copy;
        });
    };

    // ---------------- GROUPING + AGGREGATION ----------------
    const groupedData = React.useMemo(() => {

        const map = new Map<string, RecordType[]>();

        data.forEach(item => {
            if (!map.has(item.product)) {
                map.set(item.product, []);
            }
            map.get(item.product)!.push(item);
        });

        const result: GridItem[] = [];

        map.forEach((items, key) => {

            const totalQuantity = items.reduce((s, i) => s + i.quantity, 0);
            const totalAmount = items.reduce((s, i) => s + i.amount, 0);

            result.push({
                isGroup: true,
                key,
                name: key,
                count: items.length,
                totalQuantity,
                totalAmount
            });

            items.forEach(item => result.push(item));
        });

        return result;

    }, [data]);

    // ---------------- CONDITIONAL STYLE ----------------
    const getQuantityStyle = (qty: number) => ({
        color: qty < 10 ? "red" : "green",
        fontWeight: "600"
    });

    // ---------------- COLUMNS ----------------
    const columns: IColumn[] = [

        {
            key: "product",
            name: "Product",
            minWidth: 200,
            onRender: item => {

                if (item.isGroup) {
                    return <strong>{item.product}</strong>;
                }

                return (
                    <TextField
                        value={editing[item.id]?.product ?? item.product}
                        onChange={(_, v) => updateField(item.id, "product", v)}
                    />
                );
            },
            onRenderHeader: () => (
                <TextField
                    placeholder="Filter Product"
                    onChange={(_, v) => setFilter("product", v || "")}
                />
            )
        },

        {
            key: "quantity",
            name: "Quantity",
            minWidth: 100,
            onRender: item => {

                if (item.isGroup) {
                    return <strong>Total: {item.totalQuantity}</strong>;
                }

                return (
                    <TextField
                        type="number"
                        value={(editing[item.id]?.quantity ?? item.quantity).toString()}
                        onChange={(_, v) => updateField(item.id, "quantity", Number(v))}
                        styles={{
                            field: getQuantityStyle(item.quantity)
                        }}
                    />
                );
            }
        },

        {
            key: "amount",
            name: "Amount",
            minWidth: 120,
            onRender: item => {

                if (item.isGroup) {
                    return <strong>Total: {item.totalAmount}</strong>;
                }

                return (
                    <TextField
                        type="number"
                        value={(editing[item.id]?.amount ?? item.amount).toString()}
                        onChange={(_, v) => updateField(item.id, "amount", Number(v))}
                    />
                );
            }
        },

        {
            key: "actions",
            name: "Actions",
            minWidth: 100,
            onRender: item => {

                if (item.isGroup) return null;

                return (
                    <PrimaryButton
                        text="Save"
                        onClick={() => saveRow(item.id)}
                    />
                );
            }
        }
    ];

    // ---------------- PAGINATION ----------------
    const totalPages = 5; // ideally from API

    const changePage = (p: number) => {
        setRequest(prev => ({ ...prev, page: p }));
    };

    // ---------------- UI ----------------
    return (
        <Stack tokens={{ childrenGap: 10 }}>

            {loading && <Spinner label="Loading..." />}

            {/* GRID */}
            <DetailsList
                items={groupedData}
                columns={columns}
                selectionMode={SelectionMode.none}
                onShouldVirtualize={() => true}
            />

            {/* PAGINATION */}
            <Stack horizontal tokens={{ childrenGap: 10 }}>

                <DefaultButton
                    text="Prev"
                    onClick={() => changePage(request.page - 1)}
                />

                <div>Page {request.page}</div>

                <DefaultButton
                    text="Next"
                    onClick={() => changePage(request.page + 1)}
                />

            </Stack>

        </Stack>
    );
};