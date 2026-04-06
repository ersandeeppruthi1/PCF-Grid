import * as React from "react";
import {
    DetailsList,
    IColumn,
    SelectionMode,
    Stack,
    TextField,
    Spinner,
    DefaultButton,
    PrimaryButton,
    IconButton
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
    console.log("FluentGrid: Component rendered");
    console.log("FluentGrid: initialData:", initialData);
    console.log("FluentGrid: context:", context);

    // ---------------- STATE ----------------
    const [data, setData] = React.useState<RecordType[]>(initialData || []);
    const [loading, setLoading] = React.useState(false);

    console.log("FluentGrid: Current data state:", data);
    console.log("FluentGrid: Current loading state:", loading);

    const [request, setRequest] = React.useState({
        page: 1,
        pageSize: 25,
        filters: {} as Record<string, unknown>
    });

    const [editing, setEditing] = React.useState<Record<string, Partial<RecordType>>>({});
    const [expandedGroups, setExpandedGroups] = React.useState<Set<string>>(new Set());

    // Auto-expand groups when data changes (optional - for better UX)
    React.useEffect(() => {
        if (data && data.length > 0) {
            const productGroups = [...new Set(data.map(item => item.product))];
            setExpandedGroups(new Set(productGroups));
        }
    }, [data]);

    // ---------------- FETCH ----------------
    const fetchData = React.useCallback(async () => {
        console.log("FluentGrid: fetchData called");
        setLoading(true);
        try {
            console.log("FluentGrid: Making API request to /api/grid with request:", request);
            const res = await axios.post("/api/grid", request);
            console.log("FluentGrid: API response:", res.data);
            setData(res.data.items);
            console.log("FluentGrid: Data set to:", res.data.items);
        } catch (error) {
            console.error("FluentGrid: API request failed:", error);
        } finally {
            console.log("FluentGrid: Setting loading to false");
            setLoading(false);
        }
    }, [request]);

    React.useEffect(() => {
        console.log("FluentGrid: useEffect triggered for fetchData");
        console.log("FluentGrid: Current request:", request);
        const t = setTimeout(() => {
            console.log("FluentGrid: Calling fetchData after 400ms delay");
            fetchData();
        }, 400);
        return () => {
            console.log("FluentGrid: Clearing fetchData timeout");
            clearTimeout(t);
        };
    }, [fetchData]);

    // Effect to update data when initialData changes
    React.useEffect(() => {
        console.log("FluentGrid: useEffect triggered for initialData change");
        console.log("FluentGrid: New initialData:", initialData);
        if (initialData && initialData.length > 0) {
            console.log("FluentGrid: Setting data from initialData");
            setData(initialData);
        }
    }, [initialData]);

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

    // ---------------- GROUP EXPAND/COLLAPSE ----------------
    const toggleGroup = (groupKey: string) => {
        setExpandedGroups(prev => {
            const newSet = new Set(prev);
            if (newSet.has(groupKey)) {
                newSet.delete(groupKey);
            } else {
                newSet.add(groupKey);
            }
            return newSet;
        });
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
        console.log("FluentGrid: Processing groupedData with data:", data);

        const map = new Map<string, RecordType[]>();

        data.forEach(item => {
            console.log("FluentGrid: Processing item for grouping:", item);
            if (!map.has(item.product)) {
                map.set(item.product, []);
            }
            map.get(item.product)!.push(item);
        });

        console.log("FluentGrid: Grouped map:", map);

        const result: GridItem[] = [];

        map.forEach((items, key) => {
            console.log(`FluentGrid: Processing group ${key} with ${items.length} items:`, items);

            const totalQuantity = items.reduce((s, i) => s + i.quantity, 0);
            const totalAmount = items.reduce((s, i) => s + i.amount, 0);

            console.log(`FluentGrid: Group ${key} totals - quantity: ${totalQuantity}, amount: ${totalAmount}`);

            result.push({
                isGroup: true,
                key,
                name: key,
                count: items.length,
                totalQuantity,
                totalAmount
            });

            // Only add child items if the group is expanded
            if (expandedGroups.has(key)) {
                items.forEach(item => {
                    console.log(`FluentGrid: Adding item to result:`, item);
                    result.push(item);
                });
            }
        });

        console.log("FluentGrid: Final grouped result:", result);
        return result;

    }, [data, expandedGroups]);

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
                    const isExpanded = expandedGroups.has(item.key);
                    return (
                        <Stack 
                            horizontal 
                            verticalAlign="center" 
                            tokens={{ childrenGap: 8 }}
                            styles={{
                                root: {
                                    backgroundColor: '#f3f2f1',
                                    padding: '4px 8px',
                                    borderRadius: '4px',
                                    border: '1px solid #edebe9'
                                }
                            }}
                        >
                            <IconButton
                                iconProps={{ iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }}
                                onClick={() => toggleGroup(item.key)}
                                styles={{ 
                                    root: { 
                                        width: 24, 
                                        height: 24,
                                        minWidth: 24
                                    } 
                                }}
                            />
                            <strong>{item.name} ({item.count} items)</strong>
                        </Stack>
                    );
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
    console.log("FluentGrid: Rendering UI");
    console.log("FluentGrid: groupedData for render:", groupedData);
    console.log("FluentGrid: columns for render:", columns);
    console.log("FluentGrid: loading state for render:", loading);
    
    if (loading) {
        console.log("FluentGrid: Showing loading spinner");
    }
    
    console.log("FluentGrid: DetailsList will render with", groupedData.length, "items");
    
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