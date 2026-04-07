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
    IGroup,
    IGroupHeaderProps,
    GroupHeader,
    Icon,
    TooltipHost,
    Dropdown,
    IDropdownOption
} from "@fluentui/react";
import axios from "axios";

interface RecordType {
    id: string;
    product: string;
    quantity: number;
    amount: number;
    status: 'Active' | 'Inactive' | 'Pending';
    category: string;
    lastModified?: Date;
    owner?: string;
    isDirty?: boolean; // Track unsaved changes
}

interface CRMGridProps {
    data?: RecordType[];
    context?: ComponentFramework.Context<unknown>;
    onSave?: (records: RecordType[]) => Promise<void>;
    onDelete?: (recordIds: string[]) => Promise<void>;
}

export const FluentGrid: React.FC<CRMGridProps> = ({ data: initialData, context, onSave, onDelete }) => {
    console.log("CRMGrid: Component rendered");
    console.log("CRMGrid: initialData:", initialData);
    console.log("CRMGrid: context:", context);

    // ---------------- STATE ----------------
    const [data, setData] = React.useState<RecordType[]>(initialData || []);
    const [loading, setLoading] = React.useState(false);
    const [editing, setEditing] = React.useState<Record<string, Partial<RecordType>>>({});
    const [editingRows, setEditingRows] = React.useState<Set<string>>(new Set());
    const [selectedItems, setSelectedItems] = React.useState<RecordType[]>([]);
    const [dirtyRecords, setDirtyRecords] = React.useState<Set<string>>(new Set());
    const [sortBy, setSortBy] = React.useState<{ key: string; descending: boolean } | null>(null);
    const [filterText, setFilterText] = React.useState<string>('');

    console.log("CRMGrid: Current data state:", data);
    console.log("CRMGrid: Current loading state:", loading);

    // ---------------- SELECTION ----------------
    const selection = React.useMemo(() => {
        const sel = new Selection({
            onSelectionChanged: () => {
                const selected = sel.getSelection() as RecordType[];
                setSelectedItems(selected);
                console.log("Selected items:", selected);
            },
        });
        return sel;
    }, []);

    // ---------------- COMMAND BAR (CRM Style) ----------------
    const commandBarItems: ICommandBarItemProps[] = [
        {
            key: 'new',
            text: 'New',
            iconProps: { iconName: 'Add' },
            onClick: () => {
                console.log("Creating new record");
                // Add new record logic
                const newRecord: RecordType = {
                    id: `new-${Date.now()}`,
                    product: '',
                    quantity: 0,
                    amount: 0,
                    status: 'Active',
                    category: '',
                    isDirty: true
                };
                setData(prev => [newRecord, ...prev]);
                setEditingRows(prev => new Set([...prev, newRecord.id]));
                setDirtyRecords(prev => new Set([...prev, newRecord.id]));
            },
        },
        {
            key: 'edit',
            text: 'Edit',
            iconProps: { iconName: 'Edit' },
            disabled: selectedItems.length !== 1,
            onClick: () => {
                if (selectedItems.length === 1) {
                    toggleEditMode(selectedItems[0].id);
                }
            },
        },
        {
            key: 'delete',
            text: 'Delete',
            iconProps: { iconName: 'Delete' },
            disabled: selectedItems.length === 0,
            onClick: async () => {
                if (window.confirm(`Delete ${selectedItems.length} selected record(s)?`)) {
                    const selectedIds = selectedItems.map(item => item.id);
                    if (onDelete) {
                        await onDelete(selectedIds);
                    }
                    setData(prev => prev.filter(item => !selectedIds.includes(item.id)));
                    selection.setAllSelected(false);
                }
            },
        },
        {
            key: 'save',
            text: 'Save',
            iconProps: { iconName: 'Save' },
            disabled: dirtyRecords.size === 0,
            onClick: async () => {
                if (onSave) {
                    const recordsToSave = data.filter(record => dirtyRecords.has(record.id));
                    await onSave(recordsToSave);
                }
                setDirtyRecords(new Set());
                setEditingRows(new Set());
                // Mark all records as not dirty
                setData(prev => prev.map(record => ({ ...record, isDirty: false })));
            },
        },
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: () => {
                // Reset dirty state and reload
                setDirtyRecords(new Set());
                setEditingRows(new Set());
                setEditing({});
                fetchData();
            },
        }
    ];

    const commandBarFarItems: ICommandBarItemProps[] = [
        {
            key: 'export',
            text: 'Export to Excel',
            iconProps: { iconName: 'Download' },
            onClick: () => {
                const csvContent = data.map(item => 
                    `${item.product},${item.quantity},${item.amount},${item.status},${item.category}`
                ).join('\n');
                const header = 'Product,Quantity,Amount,Status,Category\n';
                const blob = new Blob([header + csvContent], { type: 'text/csv' });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'records.csv';
                a.click();
            },
        },
        {
            key: 'filter',
            text: filterText ? `Clear Filter` : 'Filter',
            iconProps: { iconName: filterText ? 'ClearFilter' : 'Filter' },
            onClick: () => {
                if (filterText) {
                    setFilterText('');
                } else {
                    // Show filter UI
                    const filter = prompt("Enter filter text:");
                    if (filter) setFilterText(filter);
                }
            },
        }
    ];

    // ---------------- FETCH DATA ----------------
    const fetchData = React.useCallback(async () => {
        console.log("CRMGrid: fetchData called");
        setLoading(true);
        try {
            console.log("CRMGrid: Making API request");
            // Simulate API call - replace with actual CRM API
            await new Promise(resolve => setTimeout(resolve, 500));
            console.log("CRMGrid: API response received");
        } catch (error) {
            console.error("CRMGrid: API request failed:", error);
        } finally {
            console.log("CRMGrid: Setting loading to false");
            setLoading(false);
        }
    }, []);

    // Effect to update data when initialData changes
    React.useEffect(() => {
        console.log("CRMGrid: useEffect triggered for initialData change");
        console.log("CRMGrid: New initialData:", initialData);
        if (initialData && initialData.length > 0) {
            console.log("CRMGrid: Setting data from initialData");
            setData(initialData);
        }
    }, [initialData]);

    // ---------------- EDIT FUNCTIONS ----------------
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
            
            // Mark record as dirty
            setDirtyRecords(prev => new Set([...prev, id]));
            
            // Update the actual data for immediate UI feedback
            setData(prev => 
                prev.map(item => 
                    item.id === id 
                        ? { ...item, [field]: value, isDirty: true }
                        : item
                )
            );
        }
    };

    // ---------------- DATA PROCESSING WITH GROUPING ----------------
    const { groups, groupedData } = React.useMemo(() => {
        let filtered = data;
        
        // Apply filtering first
        if (filterText) {
            filtered = data.filter(item => 
                Object.values(item).some(value => 
                    value?.toString().toLowerCase().includes(filterText.toLowerCase())
                )
            );
        }

        // Apply sorting
        if (sortBy) {
            filtered = [...filtered].sort((a, b) => {
                const aValue = a[sortBy.key as keyof RecordType];
                const bValue = b[sortBy.key as keyof RecordType];
                
                // Handle undefined values
                if (aValue == null && bValue == null) return 0;
                if (aValue == null) return sortBy.descending ? 1 : -1;
                if (bValue == null) return sortBy.descending ? -1 : 1;
                
                if (aValue < bValue) return sortBy.descending ? 1 : -1;
                if (aValue > bValue) return sortBy.descending ? -1 : 1;
                return 0;
            });
        }

        // Create groups by product name
        const productGroups = new Map<string, RecordType[]>();
        
        filtered.forEach(item => {
            const productName = item.product || 'Unnamed Product';
            if (!productGroups.has(productName)) {
                productGroups.set(productName, []);
            }
            productGroups.get(productName)!.push(item);
        });

        // Create FluentUI groups structure
        const groups: IGroup[] = [];
        const flatData: RecordType[] = [];
        let startIndex = 0;

        productGroups.forEach((items, productName) => {
            const totalQuantity = items.reduce((sum, item) => sum + item.quantity, 0);
            const totalAmount = items.reduce((sum, item) => sum + item.amount, 0);
            const activeCount = items.filter(item => item.status === 'Active').length;
            const inactiveCount = items.filter(item => item.status === 'Inactive').length;
            const pendingCount = items.filter(item => item.status === 'Pending').length;

            groups.push({
                key: productName,
                name: `${productName}`,
                startIndex,
                count: items.length,
                isCollapsed: false,
                level: 0,
                data: {
                    totalQuantity,
                    totalAmount,
                    activeCount,
                    inactiveCount,
                    pendingCount
                }
            });

            // Add sorted items to flat data
            flatData.push(...items);
            startIndex += items.length;
        });

        return { groups, groupedData: flatData };
    }, [data, filterText, sortBy]);

    // ---------------- STYLING ----------------
    const getStatusIcon = (status: string) => {
        switch (status) {
            case 'Active': return 'CompletedSolid';
            case 'Inactive': return 'BlockedSolid';
            case 'Pending': return 'Clock';
            default: return 'Info';
        }
    };

    const getStatusColor = (status: string) => {
        switch (status) {
            case 'Active': return '#107C10';
            case 'Inactive': return '#D13438';
            case 'Pending': return '#FF8C00';
            default: return '#605E5C';
        }
    };

    const getQuantityStyle = (qty: number) => ({
        color: qty < 10 ? "#D13438" : "#107C10",
        fontWeight: "600" as const
    });

    // ---------------- COLUMNS (CRM Style) ----------------
    const columns: IColumn[] = [
        {
            key: "product",
            name: "Product Name",
            minWidth: 200,
            maxWidth: 300,
            isResizable: true,
            isSorted: sortBy?.key === 'product',
            isSortedDescending: sortBy?.key === 'product' ? sortBy.descending : false,
            onColumnClick: () => {
                setSortBy(prev => 
                    prev?.key === 'product' 
                        ? { key: 'product', descending: !prev.descending }
                        : { key: 'product', descending: false }
                );
            },
            onRender: (item: RecordType) => {
                const isEditing = editingRows.has(item.id);
                const isDirty = dirtyRecords.has(item.id);
                
                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        {isDirty && <Icon iconName="Edit" style={{ color: '#0078D4', fontSize: 12 }} />}
                        {isEditing ? (
                            <TextField
                                value={editing[item.id]?.product ?? item.product}
                                onChange={(_, value) => updateField(item.id, "product", value)}
                                styles={{
                                    root: { width: '100%' },
                                    fieldGroup: { 
                                        height: 28,
                                        border: '1px solid #0078D4',
                                        borderRadius: 2
                                    },
                                    field: { 
                                        fontSize: 14,
                                        padding: '4px 8px'
                                    }
                                }}
                            />
                        ) : (
                            <span 
                                style={{ 
                                    fontSize: 14, 
                                    fontWeight: 500,
                                    cursor: 'pointer',
                                    color: '#323130'
                                }}
                                onClick={() => toggleEditMode(item.id)}
                                title="Click to edit"
                            >
                                {item.product || 'Unnamed Product'}
                            </span>
                        )}
                    </div>
                );
            }
        },

        {
            key: "status",
            name: "Status",
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            isSorted: sortBy?.key === 'status',
            isSortedDescending: sortBy?.key === 'status' ? sortBy.descending : false,
            onColumnClick: () => {
                setSortBy(prev => 
                    prev?.key === 'status' 
                        ? { key: 'status', descending: !prev.descending }
                        : { key: 'status', descending: false }
                );
            },
            onRender: (item: RecordType) => {
                const isEditing = editingRows.has(item.id);
                
                const statusOptions: IDropdownOption[] = [
                    { key: 'Active', text: 'Active' },
                    { key: 'Inactive', text: 'Inactive' },
                    { key: 'Pending', text: 'Pending' }
                ];

                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        {!isEditing && (
                            <Icon 
                                iconName={getStatusIcon(item.status)} 
                                style={{ 
                                    color: getStatusColor(item.status),
                                    fontSize: 12
                                }} 
                            />
                        )}
                        {isEditing ? (
                            <Dropdown
                                options={statusOptions}
                                selectedKey={editing[item.id]?.status ?? item.status}
                                onChange={(_, option) => updateField(item.id, "status", option?.key as string)}
                                styles={{
                                    root: { width: '100%' },
                                    dropdown: { 
                                        height: 28,
                                        border: '1px solid #0078D4',
                                        borderRadius: 2
                                    }
                                }}
                            />
                        ) : (
                            <span 
                                style={{ 
                                    fontSize: 14,
                                    color: getStatusColor(item.status),
                                    fontWeight: 500,
                                    cursor: 'pointer'
                                }}
                                onClick={() => toggleEditMode(item.id)}
                                title="Click to edit"
                            >
                                {item.status}
                            </span>
                        )}
                    </div>
                );
            }
        },

        {
            key: "quantity",
            name: "Quantity",
            minWidth: 100,
            maxWidth: 120,
            isResizable: true,
            isSorted: sortBy?.key === 'quantity',
            isSortedDescending: sortBy?.key === 'quantity' ? sortBy.descending : false,
            onColumnClick: () => {
                setSortBy(prev => 
                    prev?.key === 'quantity' 
                        ? { key: 'quantity', descending: !prev.descending }
                        : { key: 'quantity', descending: false }
                );
            },
            onRender: (item: RecordType) => {
                const isEditing = editingRows.has(item.id);
                
                return (
                    <div>
                        {isEditing ? (
                            <TextField
                                type="number"
                                value={(editing[item.id]?.quantity ?? item.quantity).toString()}
                                onChange={(_, value) => updateField(item.id, "quantity", Number(value))}
                                styles={{
                                    fieldGroup: { 
                                        height: 28,
                                        border: '1px solid #0078D4',
                                        borderRadius: 2
                                    },
                                    field: { 
                                        fontSize: 14,
                                        padding: '4px 8px',
                                        textAlign: 'right'
                                    }
                                }}
                            />
                        ) : (
                            <span 
                                style={{
                                    ...getQuantityStyle(item.quantity), 
                                    cursor: 'pointer', 
                                    fontSize: 14,
                                    display: 'block',
                                    textAlign: 'right'
                                }}
                                onClick={() => toggleEditMode(item.id)}
                                title="Click to edit"
                            >
                                {item.quantity.toLocaleString()}
                            </span>
                        )}
                    </div>
                );
            }
        },

        {
            key: "amount",
            name: "Amount",
            minWidth: 120,
            maxWidth: 150,
            isResizable: true,
            isSorted: sortBy?.key === 'amount',
            isSortedDescending: sortBy?.key === 'amount' ? sortBy.descending : false,
            onColumnClick: () => {
                setSortBy(prev => 
                    prev?.key === 'amount' 
                        ? { key: 'amount', descending: !prev.descending }
                        : { key: 'amount', descending: false }
                );
            },
            onRender: (item: RecordType) => {
                const isEditing = editingRows.has(item.id);
                
                return (
                    <div>
                        {isEditing ? (
                            <TextField
                                type="number"
                                value={(editing[item.id]?.amount ?? item.amount).toString()}
                                onChange={(_, value) => updateField(item.id, "amount", Number(value))}
                                prefix="$"
                                styles={{
                                    fieldGroup: { 
                                        height: 28,
                                        border: '1px solid #0078D4',
                                        borderRadius: 2
                                    },
                                    field: { 
                                        fontSize: 14,
                                        padding: '4px 8px',
                                        textAlign: 'right'
                                    }
                                }}
                            />
                        ) : (
                            <span 
                                onClick={() => toggleEditMode(item.id)}
                                style={{
                                    cursor: 'pointer', 
                                    fontWeight: 500, 
                                    fontSize: 14,
                                    color: '#323130',
                                    display: 'block',
                                    textAlign: 'right'
                                }}
                                title="Click to edit"
                            >
                                ${item.amount.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                            </span>
                        )}
                    </div>
                );
            }
        },

        {
            key: "category",
            name: "Category",
            minWidth: 150,
            maxWidth: 200,
            isResizable: true,
            isSorted: sortBy?.key === 'category',
            isSortedDescending: sortBy?.key === 'category' ? sortBy.descending : false,
            onColumnClick: () => {
                setSortBy(prev => 
                    prev?.key === 'category' 
                        ? { key: 'category', descending: !prev.descending }
                        : { key: 'category', descending: false }
                );
            },
            onRender: (item: RecordType) => {
                const isEditing = editingRows.has(item.id);
                
                return (
                    <div>
                        {isEditing ? (
                            <TextField
                                value={editing[item.id]?.category ?? item.category}
                                onChange={(_, value) => updateField(item.id, "category", value)}
                                styles={{
                                    fieldGroup: { 
                                        height: 28,
                                        border: '1px solid #0078D4',
                                        borderRadius: 2
                                    },
                                    field: { 
                                        fontSize: 14,
                                        padding: '4px 8px'
                                    }
                                }}
                            />
                        ) : (
                            <span 
                                style={{ 
                                    fontSize: 14,
                                    color: '#323130',
                                    cursor: 'pointer'
                                }}
                                onClick={() => toggleEditMode(item.id)}
                                title="Click to edit"
                            >
                                {item.category || 'Uncategorized'}
                            </span>
                        )}
                    </div>
                );
            }
        }
    ];

    // ---------------- UI ----------------
    console.log("CRMGrid: Rendering UI");
    console.log("CRMGrid: groupedData for render:", groupedData);
    console.log("CRMGrid: groups for render:", groups);
    console.log("CRMGrid: columns for render:", columns);
    console.log("CRMGrid: loading state for render:", loading);
    
    if (loading) {
        console.log("CRMGrid: Showing loading spinner");
    }
    
    console.log("CRMGrid: DetailsList will render with", groupedData.length, "items in", groups.length, "groups");
    
    return (
        <div style={{ 
            height: '100vh', 
            display: 'flex', 
            flexDirection: 'column',
            backgroundColor: '#f3f2f1' 
        }}>
            
            {/* COMMAND BAR */}
            <CommandBar
                items={commandBarItems}
                farItems={commandBarFarItems}
                styles={{
                    root: {
                        backgroundColor: 'white',
                        borderBottom: '1px solid #edebe9',
                        padding: '0 16px',
                        boxShadow: '0 1px 2px rgba(0, 0, 0, 0.1)'
                    }
                }}
            />

            {loading && (
                <div style={{ 
                    padding: '20px',
                    textAlign: 'center',
                    backgroundColor: 'white' 
                }}>
                    <Spinner label="Loading records..." />
                </div>
            )}

            {/* DIRTY RECORDS WARNING */}
            {dirtyRecords.size > 0 && (
                <div style={{
                    backgroundColor: '#fff4ce',
                    padding: '8px 16px',
                    borderBottom: '1px solid #edebe9',
                    display: 'flex',
                    alignItems: 'center',
                    gap: 8
                }}>
                    <Icon iconName="Warning" style={{ color: '#797673' }} />
                    <span style={{ fontSize: 14, color: '#323130' }}>
                        {dirtyRecords.size} record{dirtyRecords.size !== 1 ? 's' : ''} have unsaved changes
                    </span>
                </div>
            )}

            {/* SELECTION INFO */}
            {selectedItems.length > 0 && (
                <div style={{
                    backgroundColor: '#deecf9',
                    padding: '8px 16px',
                    borderBottom: '1px solid #edebe9',
                    fontSize: '14px',
                    color: '#323130',
                    display: 'flex',
                    alignItems: 'center',
                    gap: 8
                }}>
                    <Icon iconName="CheckboxComposite" style={{ color: '#0078d4' }} />
                    <span>
                        {selectedItems.length} record{selectedItems.length !== 1 ? 's' : ''} selected
                    </span>
                </div>
            )}

            {/* FILTER INFO */}
            {filterText && (
                <div style={{
                    backgroundColor: '#e1dfdd',
                    padding: '6px 16px',
                    borderBottom: '1px solid #edebe9',
                    fontSize: '12px',
                    color: '#605e5c',
                    display: 'flex',
                    alignItems: 'center',
                    gap: 8
                }}>
                    <Icon iconName="Filter" style={{ fontSize: 12 }} />
                    <span>Filtered by: "{filterText}" • {groupedData.length} of {data.length} records shown in {groups.length} groups</span>
                </div>
            )}

            {/* MAIN GRID AREA */}
            <div style={{ 
                flex: 1,
                backgroundColor: 'white',
                border: '1px solid #edebe9',
                margin: '8px 16px',
                borderRadius: 2,
                overflow: 'hidden',
                display: 'flex',
                flexDirection: 'column'
            }}>
                <DetailsList
                    items={groupedData}
                    columns={columns}
                    groups={groups}
                    selectionMode={SelectionMode.multiple}
                    selection={selection}
                    checkboxVisibility={CheckboxVisibility.always}
                    layoutMode={DetailsListLayoutMode.justified}
                    isHeaderVisible={true}
                    compact={false} // CRM grids are not compact
                    groupProps={{
                        onRenderHeader: (props?: IGroupHeaderProps) => {
                            if (!props || !props.group) return null;
                            
                            const group = props.group;
                            const groupData = group.data || {};
                            const { totalQuantity, totalAmount, activeCount, inactiveCount, pendingCount } = groupData;

                            return (
                                <div
                                    onClick={() => props.onToggleCollapse?.(group)}
                                    style={{
                                        padding: '8px 12px',
                                        backgroundColor: '#f8f7f6',
                                        borderBottom: '1px solid #edebe9',
                                        display: 'flex',
                                        alignItems: 'center',
                                        gap: 12,
                                        fontWeight: 600,
                                        fontSize: 14,
                                        color: '#323130',
                                        cursor: 'pointer'
                                    }}
                                >
                                    <Icon 
                                        iconName={group.isCollapsed ? 'ChevronRight' : 'ChevronDown'} 
                                        style={{ fontSize: 12, color: '#605e5c' }} 
                                    />
                                    <Icon iconName="Product" style={{ fontSize: 14, color: '#0078d4' }} />
                                    <span style={{ fontWeight: 600 }}>
                                        {group.name} ({group.count} records)
                                    </span>
                                    <div style={{ marginLeft: 'auto', display: 'flex', gap: 16, fontSize: 12, color: '#605e5c' }}>
                                        {totalQuantity !== undefined && (
                                            <span>Total Qty: <strong>{totalQuantity.toLocaleString()}</strong></span>
                                        )}
                                        {totalAmount !== undefined && (
                                            <span>Total Amount: <strong>${totalAmount.toLocaleString()}</strong></span>
                                        )}
                                        {activeCount !== undefined && (
                                            <span style={{ color: '#107C10' }}>Active: <strong>{activeCount}</strong></span>
                                        )}
                                        {pendingCount !== undefined && pendingCount > 0 && (
                                            <span style={{ color: '#FF8C00' }}>Pending: <strong>{pendingCount}</strong></span>
                                        )}
                                        {inactiveCount !== undefined && inactiveCount > 0 && (
                                            <span style={{ color: '#D13438' }}>Inactive: <strong>{inactiveCount}</strong></span>
                                        )}
                                    </div>
                                </div>
                            );
                        }
                    }}
                    styles={{
                        root: {
                            flex: 1,
                            '& .ms-DetailsHeader': {
                                backgroundColor: '#faf9f8',
                                borderBottom: '1px solid #edebe9',
                                paddingTop: 0,
                                height: 42
                            },
                            '& .ms-DetailsHeader-cell': {
                                fontSize: 14,
                                fontWeight: 600,
                                color: '#323130',
                                padding: '0 12px',
                                '&:hover': {
                                    backgroundColor: '#f3f2f1'
                                },
                                '&:active': {
                                    backgroundColor: '#edebe9'
                                }
                            },
                            '& .ms-DetailsHeader-cellName': {
                                fontSize: 14,
                                fontWeight: 600
                            },
                            '& .ms-DetailsRow': {
                                borderBottom: '1px solid #f3f2f1',
                                minHeight: 40,
                                '&:hover': {
                                    backgroundColor: '#f8f7f6'
                                },
                                '&.is-selected': {
                                    backgroundColor: '#deecf9',
                                    '&:hover': {
                                        backgroundColor: '#cce1f5'
                                    }
                                },
                                '&.is-selected:after': {
                                    borderLeft: '3px solid #0078d4'
                                }
                            },
                            '& .ms-DetailsRow-cell': {
                                fontSize: 14,
                                padding: '8px 12px',
                                minHeight: 40,
                                display: 'flex',
                                alignItems: 'center',
                                color: '#323130'
                            },
                            '& .ms-Check': {
                                height: 20,
                                width: 20
                            },
                            '& .ms-CheckBox': {
                                marginRight: 8
                            }
                        }
                    }}
                />

                {/* FOOTER */}
                <div style={{
                    backgroundColor: '#faf9f8',
                    borderTop: '1px solid #edebe9',
                    padding: '8px 16px',
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    fontSize: 14,
                    color: '#605e5c',
                    minHeight: 44
                }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <Icon iconName="Documentation" style={{ fontSize: 16 }} />
                        <span>
                            {groupedData.length} record{groupedData.length !== 1 ? 's' : ''} in {groups.length} group{groups.length !== 1 ? 's' : ''}
                            {filterText && ` (filtered from ${data.length})`}
                        </span>
                    </div>
                    
                    {dirtyRecords.size > 0 && (
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                            <Icon iconName="EditNote" style={{ fontSize: 16, color: '#0078d4' }} />
                            <span style={{ color: '#0078d4' }}>
                                {dirtyRecords.size} unsaved change{dirtyRecords.size !== 1 ? 's' : ''}
                            </span>
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
};