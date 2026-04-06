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
    IconButton,
    CommandBar,
    ICommandBarItemProps,
    Selection,
    CheckboxVisibility,
    DetailsListLayoutMode,
    DetailsRow,
    IDetailsRowProps
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
    const [editingRows, setEditingRows] = React.useState<Set<string>>(new Set());
    const [expandedGroups, setExpandedGroups] = React.useState<Set<string>>(new Set());
    const [selectedItems, setSelectedItems] = React.useState<RecordType[]>([]);

    // Auto-expand groups when data changes (optional - for better UX)
    React.useEffect(() => {
        if (data && data.length > 0) {
            const productGroups = [...new Set(data.map(item => item.product))];
            setExpandedGroups(new Set(productGroups));
        }
    }, [data]);

    // ---------------- SELECTION ----------------
    const selection = React.useMemo(() => {
        const sel = new Selection({
            onSelectionChanged: () => {
                const selected = sel.getSelection() as GridItem[];
                // Filter out group items and only include actual records
                const selectedRecords = selected.filter((item: GridItem) => !('isGroup' in item)) as RecordType[];
                setSelectedItems(selectedRecords);
                console.log("Selected items:", selectedRecords);
            },
        });
        return sel;
    }, []);

    // ---------------- COMMAND BAR ----------------
    const commandBarItems: ICommandBarItemProps[] = [
        {
            key: 'edit',
            text: 'Edit Selected',
            iconProps: { iconName: 'Edit' },
            disabled: selectedItems.length === 0,
            onClick: () => {
                console.log("Editing selected items:", selectedItems);
                selectedItems.forEach((item: RecordType) => toggleEditMode(item.id));
            },
        },
        {
            key: 'delete',
            text: 'Delete Selected',
            iconProps: { iconName: 'Delete' },
            disabled: selectedItems.length === 0,
            onClick: () => {
                console.log('Delete selected items:', selectedItems);
                // Add your delete logic here
                if (window.confirm(`Are you sure you want to delete ${selectedItems.length} selected item(s)?`)) {
                    // Remove selected items from data
                    const selectedIds = selectedItems.map(item => item.id);
                    setData(prev => prev.filter(item => !selectedIds.includes(item.id)));
                    selection.setAllSelected(false);
                }
            },
        },
        {
            key: 'export',
            text: 'Export Selected',
            iconProps: { iconName: 'Download' },
            disabled: selectedItems.length === 0,
            onClick: () => {
                console.log('Export selected items:', selectedItems);
                // Add your export logic here
                const csvContent = selectedItems.map(item => 
                    `${item.product},${item.quantity},${item.amount}`
                ).join('\n');
                const header = 'Product,Quantity,Amount\n';
                const blob = new Blob([header + csvContent], { type: 'text/csv' });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'selected_items.csv';
                a.click();
            },
        },
    ];

    const commandBarFarItems: ICommandBarItemProps[] = [
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: () => fetchData(),
        },
        {
            key: 'selectAll',
            text: selectedItems.length === data.length ? 'Clear Selection' : 'Select All',
            iconProps: { iconName: selectedItems.length === data.length ? 'Clear' : 'SelectAll' },
            onClick: () => {
                if (selectedItems.length === data.length) {
                    selection.setAllSelected(false);
                } else {
                    selection.setAllSelected(true);
                }
            },
        },
    ];

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
    const toggleEditMode = (id: string) => {
        setEditingRows(prev => {
            const newSet = new Set(prev);
            if (newSet.has(id)) {
                newSet.delete(id);
            } else {
                newSet.add(id);
            }
            return newSet;
        });
    };

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

        // Exit edit mode after saving
        setEditingRows(prev => {
            const newSet = new Set(prev);
            newSet.delete(id);
            return newSet;
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
        fontWeight: "600" as const
    });

    // ---------------- CUSTOM ROW RENDERER ----------------
    const onRenderRow = (props: IDetailsRowProps | undefined) => {
        if (!props) return null;

        const item = props.item as GridItem;

        let backgroundColor = '';

        if (!('isGroup' in item)) {
            const record = item as RecordType;
            backgroundColor = record.quantity < 10 ? "#ffebee" : "#e8f5e8";
        }

        return (
            <DetailsRow
                {...props}
                styles={{
                    root: {
                        backgroundColor,
                        minHeight: 28
                    }
                }}
            />
        );
    };

    // ---------------- COLUMNS ----------------
    const columns: IColumn[] = [

        {
            key: "product",
            name: "Product",
            minWidth: 200,
            isResizable: true,
            onRender: (item: GridItem) => {
                if ('isGroup' in item && item.isGroup) {
                    const isExpanded = expandedGroups.has(item.key);
                    return (
                        <Stack 
                            horizontal 
                            verticalAlign="center" 
                            tokens={{ childrenGap: 8 }}
                            styles={{
                                root: {
                                    backgroundColor: '#f8f9fa',
                                    padding: '4px 8px',
                                    borderRadius: '2px',
                                    border: '1px solid #dee2e6',
                                    fontWeight: 600,
                                    color: '#495057'
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

                const recordItem = item as RecordType;
                // Product name is read-only, not editable
                return (
                    <span style={{ 
                        fontSize: 13, 
                        fontWeight: 500,
                        paddingLeft: '32px', // Indent child items under group
                        display: 'block'
                    }}>
                        {recordItem.product}
                    </span>
                );
            },
            onRenderHeader: () => (
                <Stack>
                    <span style={{ fontWeight: 600, marginBottom: 4 }}>Product</span>
                    <TextField
                        placeholder="Filter Product"
                        onChange={(_, v) => setFilter("product", v || "")}
                        styles={{
                            root: { marginTop: 4 },
                            field: { fontSize: 12 }
                        }}
                    />
                </Stack>
            )
        },

        {
            key: "quantity",
            name: "Quantity",
            minWidth: 100,
            isResizable: true,
            onRender: (item: GridItem) => {
                if ('isGroup' in item && item.isGroup) {
                    return (
                        <div style={{ fontWeight: 600, color: '#495057' }}>
                            <strong>Total: {item.totalQuantity}</strong>
                        </div>
                    );
                }

                const recordItem = item as RecordType;
                const isEditing = editingRows.has(recordItem.id);
                
                if (isEditing) {
                    return (
                        <TextField
                            type="number"
                            value={(editing[recordItem.id]?.quantity ?? recordItem.quantity).toString()}
                            onChange={(_, v) => updateField(recordItem.id, "quantity", Number(v))}
                            styles={{
                                field: {
                                    ...getQuantityStyle(recordItem.quantity),
                                    fontSize: 13,
                                    minHeight: 20,
                                    padding: '2px 4px'
                                },
                                root: { height: 24 }
                            }}
                        />
                    );
                } else {
                    return (
                        <span 
                            style={{...getQuantityStyle(recordItem.quantity), cursor: 'pointer', fontSize: 13}}
                            onClick={() => toggleEditMode(recordItem.id)}
                            role="button"
                            tabIndex={0}
                            onKeyDown={(e) => e.key === 'Enter' && toggleEditMode(recordItem.id)}
                        >
                            {recordItem.quantity}
                        </span>
                    );
                }
            }
        },

        {
            key: "amount",
            name: "Amount",
            minWidth: 120,
            isResizable: true,
            onRender: (item: GridItem) => {
                if ('isGroup' in item && item.isGroup) {
                    return (
                        <div style={{ fontWeight: 600, color: '#495057' }}>
                            <strong>Total: ${item.totalAmount.toFixed(2)}</strong>
                        </div>
                    );
                }

                const recordItem = item as RecordType;
                const isEditing = editingRows.has(recordItem.id);
                
                if (isEditing) {
                    return (
                        <TextField
                            type="number"
                            value={(editing[recordItem.id]?.amount ?? recordItem.amount).toString()}
                            onChange={(_, v) => updateField(recordItem.id, "amount", Number(v))}
                            styles={{
                                field: { 
                                    fontWeight: 500,
                                    fontSize: 13,
                                    minHeight: 20,
                                    padding: '2px 4px'
                                },
                                root: { height: 24 }
                            }}
                        />
                    );
                } else {
                    return (
                        <span 
                            onClick={() => toggleEditMode(recordItem.id)}
                            role="button"
                            tabIndex={0}
                            onKeyDown={(e) => e.key === 'Enter' && toggleEditMode(recordItem.id)}
                            style={{cursor: 'pointer', fontWeight: 500, fontSize: 13}}
                        >
                            ${recordItem.amount.toFixed(2)}
                        </span>
                    );
                }
            }
        },

        {
            key: "actions",
            name: "Actions",
            minWidth: 120,
            isResizable: false,
            onRender: (item: GridItem) => {
                if ('isGroup' in item && item.isGroup) return null;

                const recordItem = item as RecordType;
                const isEditing = editingRows.has(recordItem.id);

                if (isEditing) {
                    return (
                        <Stack horizontal tokens={{ childrenGap: 8 }}>
                            <PrimaryButton
                                text="Save"
                                onClick={() => saveRow(recordItem.id)}
                                styles={{
                                    root: { minWidth: 40, height: 20, fontSize: 11 }
                                }}
                            />
                            <DefaultButton
                                text="Cancel"
                                onClick={() => toggleEditMode(recordItem.id)}
                                styles={{
                                    root: { minWidth: 40, height: 20, fontSize: 11 }
                                }}
                            />
                        </Stack>
                    );
                } else {
                    return (
                        <DefaultButton
                            text="Edit"
                            onClick={() => toggleEditMode(recordItem.id)}
                            styles={{
                                root: { minWidth: 40, height: 20, fontSize: 11 }
                            }}
                        />
                    );
                }
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
        <Stack tokens={{ childrenGap: 0 }}>
            
            {/* COMMAND BAR */}
            <CommandBar
                items={commandBarItems}
                farItems={commandBarFarItems}
                styles={{
                    root: {
                        backgroundColor: '#f8f9fa',
                        borderBottom: '1px solid #dee2e6',
                        padding: '0 16px'
                    }
                }}
            />

            {loading && (
                <Stack horizontalAlign="center" styles={{ root: { padding: 20 } }}>
                    <Spinner label="Loading..." />
                </Stack>
            )}

            {/* SELECTION INFO */}
            {selectedItems.length > 0 && (
                <div style={{
                    backgroundColor: '#e3f2fd',
                    padding: '8px 16px',
                    borderBottom: '1px solid #dee2e6',
                    fontSize: '14px',
                    color: '#1976d2'
                }}>
                    {selectedItems.length} item{selectedItems.length !== 1 ? 's' : ''} selected
                </div>
            )}

            {/* GRID */}
            <div style={{ 
                border: '1px solid #dee2e6', 
                borderTop: selectedItems.length > 0 ? 'none' : '1px solid #dee2e6',
                backgroundColor: 'white'
            }}>
                <DetailsList
                    items={groupedData}
                    columns={columns}
                    selectionMode={SelectionMode.multiple}
                    selection={selection}
                    checkboxVisibility={CheckboxVisibility.always}
                    layoutMode={DetailsListLayoutMode.justified}
                    isHeaderVisible={true}
                    onShouldVirtualize={() => false}
                    onRenderRow={onRenderRow}
                    styles={{
                        root: {
                            '& .ms-DetailsHeader': {
                                backgroundColor: '#f8f9fa',
                                borderBottom: '2px solid #dee2e6',
                                minHeight: '32px'
                            },
                            '& .ms-DetailsHeader-cell': {
                                fontSize: 13,
                                fontWeight: 600,
                                color: '#495057',
                                padding: '4px 8px'
                            },
                            '& .ms-DetailsRow': {
                                borderBottom: '1px solid #ededed',
                                minHeight: '28px',
                                '&:hover': {
                                    filter: 'brightness(0.95)'
                                }
                            },
                            '& .ms-DetailsRow-cell': {
                                fontSize: 13,
                                padding: '2px 8px',
                                minHeight: '28px',
                                height: '28px',
                                display: 'flex',
                                alignItems: 'center'
                            },
                            '& .ms-Check': {
                                height: '18px',
                                width: '18px'
                            }
                        }
                    }}
                />
            </div>

            {/* PAGINATION */}
            <Stack 
                horizontal 
                horizontalAlign="space-between" 
                verticalAlign="center"
                tokens={{ childrenGap: 10 }}
                styles={{
                    root: {
                        backgroundColor: '#f8f9fa',
                        borderTop: '1px solid #dee2e6',
                        padding: '12px 16px'
                    }
                }}
            >
                <Stack horizontal tokens={{ childrenGap: 10 }}>
                    <DefaultButton
                        text="Previous"
                        iconProps={{ iconName: 'ChevronLeft' }}
                        onClick={() => changePage(request.page - 1)}
                        disabled={request.page <= 1}
                    />
                    <DefaultButton
                        text="Next"
                        iconProps={{ iconName: 'ChevronRight' }}
                        onClick={() => changePage(request.page + 1)}
                        disabled={request.page >= totalPages}
                    />
                </Stack>
                
                <span style={{ fontSize: 14, color: '#495057' }}>
                    Page {request.page} of {totalPages}
                </span>
                
                <span style={{ fontSize: 14, color: '#6c757d' }}>
                    Total: {data.length} items
                </span>
            </Stack>

        </Stack>
    );
};