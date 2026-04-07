import * as React from "react";
import {
    DetailsList,
    IColumn,
    SelectionMode,
    TextField,
    Spinner,
    CommandBar,
    ICommandBarItemProps,
    Selection,
    CheckboxVisibility,
    DetailsListLayoutMode,
    ConstrainMode,
    IGroup,
    IGroupHeaderProps,
    GroupHeader,
    Icon,
    Dropdown,
    IDropdownOption
} from "@fluentui/react";

interface RecordType {
    id: string;
    product: string;
    quantity: number;
    amount: number;
    status: 'Open' | 'Active' | 'Resolved';
    category: string;
    isDirty?: boolean;
}

interface CRMGridProps {
    data?: RecordType[];
}

export const FluentGrid: React.FC<CRMGridProps> = ({ data: initialData }) => {

    // ---------------- STATE ----------------
    const [data, setData] = React.useState<RecordType[]>(initialData || []);
    const [editingRows, setEditingRows] = React.useState<Set<string>>(new Set());
    const [dirtyRecords, setDirtyRecords] = React.useState<Set<string>>(new Set());
    const [selectedItems, setSelectedItems] = React.useState<RecordType[]>([]);
    const [filterText, setFilterText] = React.useState("");

    // Add CSS to override FluentUI background colors
    React.useEffect(() => {
        const styleId = 'fluent-grid-row-colors';
        let styleElement = document.getElementById(styleId) as HTMLStyleElement;
        
        if (!styleElement) {
            styleElement = document.createElement('style');
            styleElement.id = styleId;
            document.head.appendChild(styleElement);
        }
        
        styleElement.innerHTML = `
            .ms-DetailsList [class*="root-"] {
                background-color: transparent !important;
            }
            .ms-DetailsList .ms-DetailsRow {
                background-color: transparent !important;
            }
            .fluent-grid-row-red {
                background-color: #ffcccb !important;
            }
            .fluent-grid-row-green {
                background-color: #90ee90 !important;
            }
        `;
        
        return () => {
            if (styleElement && styleElement.parentNode) {
                styleElement.parentNode.removeChild(styleElement);
            }
        };
    }, []);

    // ---------------- SELECTION ----------------
    const selectionRef = React.useRef(
        new Selection({
            onSelectionChanged: () => {
                const selected = selectionRef.current.getSelection() as RecordType[];
                setSelectedItems(selected);
            }
        })
    );

    // ---------------- COMMAND BAR ----------------
    const commandBarItems: ICommandBarItemProps[] = [
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: async () => {
                try {
                    console.log("Refreshing grid data from server...");
                    
                    // Show loading state (optional)
                    // setLoading(true);
                    
                    // Dynamics 365 Web API call to fetch fresh data
                    const entityName = "salesorderdetails"; // Replace with your actual entity logical name
                    const baseUrl = "https://orgc6359062.crm8.dynamics.com";
                    const apiUrl = `${baseUrl}/api/data/v9.2/${entityName}?$select=salesorderdetailid,quantity,extendedamount,productname,_new_categoryid_value&$expand=new_categoryid($select=title)`;
                    
                    const response = await fetch(apiUrl, {
                        method: 'GET',
                        headers: {
                            'Accept': 'application/json',
                            'OData-MaxVersion': '4.0',
                            'OData-Version': '4.0',
                            'Prefer': 'odata.include-annotations="*"'
                        }
                    });
                    
                    if (!response.ok) {
                        const errorText = await response.text();
                        throw new Error(`Failed to fetch data: ${response.status} - ${errorText}`);
                    }
                    
                    const result = await response.json();
                    
                    // Define interface for server response
                    interface ServerRecord {
                        salesorderdetailid: string;
                        quantity?: number;
                        extendedamount?: number;
                        productname?: string;
                        _new_categoryid_value?: string;
                        new_categoryid?: {
                            title: string;
                        };
                    }
                    
                    // Transform server data to our RecordType format
                    const refreshedData: RecordType[] = result.value.map((item: ServerRecord) => ({
                        id: item.salesorderdetailid,
                        product: item.productname || "Unknown Product",
                        quantity: item.quantity || 0,
                        amount: item.extendedamount || 0,
                        status: 'Active', // Default status since statuscode field doesn't exist
                        category: item.new_categoryid?.title || "General", // Get category title from expanded lookup
                        isDirty: false
                    }));
                    
                    // Update the grid with fresh server data
                    setData(refreshedData);
                    
                    // Clear all editing states
                    setEditingRows(new Set());
                    setDirtyRecords(new Set());
                    setSelectedItems([]);
                    setFilterText("");
                    
                    console.log(`Successfully refreshed ${refreshedData.length} records from server`);
                    
                } catch (error) {
                    console.error("Error refreshing data from server:", error);
                    
                    // Fallback to initial data if server call fails
                    if (initialData) {
                        console.log("Falling back to initial data");
                        setData(initialData);
                        setEditingRows(new Set());
                        setDirtyRecords(new Set());
                        setSelectedItems([]);
                        setFilterText("");
                    }
                    
                    alert("Failed to refresh data from server. Please check your connection and try again.");
                }
            }
        },
        // Show Update and Cancel buttons when there are editing rows
        ...(editingRows.size > 0 ? [
            {
                key: 'update',
                text: 'Update',
                iconProps: { iconName: 'Save' },
                onClick: async () => {
                    // Save changes for all editing rows to server
                    console.log("Updating records:", Array.from(editingRows));
                    
                    try {
                        // Get all dirty records that need to be updated
                        const recordsToUpdate = data.filter(record => 
                            editingRows.has(record.id) || dirtyRecords.has(record.id)
                        );
                        
                        // Update each record on the server
                        const updatePromises = recordsToUpdate.map(async (record) => {
                            try {
                                // Prepare data for server update
                                const updateData = {
                                    id: record.id,
                                    product: record.product,
                                    quantity: record.quantity,
                                    amount: record.amount,
                                    status: record.status,
                                    category: record.category
                                };
                                
                                // Simulate server API call - replace with actual API endpoint
                                console.log("Saving to server:", updateData);
                                
                                // Dynamics 365 Web API call - using proper entity endpoint
                                const entityName = "salesorderdetails"; // Replace with your actual entity logical name
                                const baseUrl = "https://orgc6359062.crm8.dynamics.com";
                                const apiUrl = `${baseUrl}/api/data/v9.2/${entityName}(${record.id})`;
                                
                                const response = await fetch(apiUrl, {
                                    method: 'PATCH', // Use PATCH for updates in Dynamics 365
                                    headers: {
                                        'Content-Type': 'application/json',
                                        'Accept': 'application/json',
                                        'OData-MaxVersion': '4.0',
                                        'OData-Version': '4.0',
                                        'If-Match': '*' // Optimistic concurrency
                                    },
                                    body: JSON.stringify({
                                        quantity: updateData.quantity,
                                        extendedamount: updateData.amount
                                        // Map other fields to Dynamics 365 field names as needed
                                        // productname: updateData.product,
                                        // statuscode: updateData.status
                                    })
                                });
                                
                                if (!response.ok) {
                                    const errorText = await response.text();
                                    throw new Error(`Failed to update record ${record.id}: ${response.status} - ${errorText}`);
                                }
                                
                                return { success: true, id: record.id };
                                
                            } catch (error) {
                                console.error(`Error updating record ${record.id}:`, error);
                                return { success: false, id: record.id, error };
                            }
                        });
                        
                        // Wait for all updates to complete
                        const results = await Promise.allSettled(updatePromises);
                        
                        // Check results and handle any failures
                        const failures = results
                            .filter(result => result.status === 'rejected' || 
                                    (result.status === 'fulfilled' && !result.value.success))
                            .map(result => result.status === 'fulfilled' ? result.value : result.reason);
                        
                        if (failures.length > 0) {
                            console.error("Some updates failed:", failures);
                            alert(`Failed to update ${failures.length} record(s). Please try again.`);
                            return;
                        }
                        
                        // Mark records as clean (no longer dirty)
                        setData(prev => 
                            prev.map(record => 
                                editingRows.has(record.id) || dirtyRecords.has(record.id)
                                    ? { ...record, isDirty: false }
                                    : record
                            )
                        );
                        
                        // Clear editing and dirty states
                        setEditingRows(new Set());
                        setDirtyRecords(new Set());
                        
                        console.log("All records updated successfully");
                        
                        // Optional: Show success message
                        // alert("Records updated successfully!");
                        
                    } catch (error) {
                        console.error("Error during update process:", error);
                        alert("Failed to update records. Please try again.");
                    }
                }
            },
            {
                key: 'cancel',
                text: 'Cancel',
                iconProps: { iconName: 'Cancel' },
                onClick: () => {
                    // Cancel changes - restore original values
                    if (initialData) {
                        const originalData = new Map(initialData.map(item => [item.id, item]));
                        setData(prev => 
                            prev.map(item => {
                                if (editingRows.has(item.id) && originalData.has(item.id)) {
                                    return originalData.get(item.id)!;
                                }
                                return item;
                            })
                        );
                    }
                    setEditingRows(new Set());
                    setDirtyRecords(new Set());
                }
            }
        ] : [])
    ];

    // ---------------- EDIT ----------------
    const toggleEdit = (id: string) => {
        setEditingRows(prev => {
            const s = new Set(prev);
            if (s.has(id)) {
                s.delete(id);
            } else {
                s.add(id);
            }
            return s;
        });
    };

    const updateField = (id: string, field: keyof RecordType, value: string | number) => {
        setDirtyRecords(prev => new Set([...prev, id]));

        setData(prev =>
            prev.map(r =>
                r.id === id ? { ...r, [field]: value, isDirty: true } : r
            )
        );
    };

    // ---------------- GROUPING ----------------
    const { groupedData, groups } = React.useMemo(() => {

        let filtered = data;

        if (filterText) {
            filtered = data.filter(d =>
                Object.values(d).some(v =>
                    v?.toString().toLowerCase().includes(filterText.toLowerCase())
                )
            );
        }

        const map = new Map<string, RecordType[]>();

        filtered.forEach(item => {
            const key = item.category || "Uncategorized";
            if (!map.has(key)) map.set(key, []);
            map.get(key)!.push(item);
        });

        const groups: IGroup[] = [];
        const flat: RecordType[] = [];
        let index = 0;

        map.forEach((items, key) => {
            const totalQty = items.reduce((a, b) => a + b.quantity, 0);
            const totalAmt = items.reduce((a, b) => a + b.amount, 0);

            groups.push({
                key,
                name: key,
                startIndex: index,
                count: items.length,
                level: 0,
                data: {
                    totalQty,
                    totalAmt
                }
            });

            flat.push(...items);
            index += items.length;
        });

        return { groupedData: flat, groups };

    }, [data, filterText]);

    // Update selection items when data changes
    React.useEffect(() => {
        if (selectionRef.current && groupedData) {
            selectionRef.current.setItems(groupedData.map(item => ({ ...item, key: item.id })), true);
        }
    }, [groupedData]);

    // ---------------- COLUMNS ----------------
    const columns: IColumn[] = [
        {
            key: "product",
            name: "Product",
            fieldName: "product",
            minWidth: 275,
            maxWidth: 275,
            isResizable: false,
            onRender: (item: RecordType) => (
                <div style={{ width: '259px', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                    <span style={{ cursor: 'default' }}>
                        {item.product || "Unnamed"}
                    </span>
                </div>
            )
        },
        {
            key: "quantity",
            name: "Qty",
            fieldName: "quantity",
            minWidth: 80,
            maxWidth: 80,
            isResizable: false,
            onRender: (item: RecordType) => (
                <div style={{ width: '80px', textAlign: 'right' }}>
                    {editingRows.has(item.id) ? (
                        <TextField
                            type="number"
                            value={item.quantity.toString()}
                            onChange={(_, v) => updateField(item.id, "quantity", Number(v || "0"))}
                            styles={{ root: { width: '100%' } }}
                        />
                    ) : (
                        <span onClick={() => toggleEdit(item.id)} style={{ cursor: 'pointer', paddingRight: '65px' }}>
                            {item.quantity}
                        </span>
                    )}
                </div>
            )
        },
        {
            key: "amount",
            name: "Amount",
            fieldName: "amount",
            minWidth: 100,
            maxWidth: 100,
            isResizable: false,
            onRender: (item: RecordType) => (
                <div style={{ width: '100px', textAlign: 'right' }}>
                    <span style={{ paddingRight: '47px' }}>${item.amount.toFixed(2)}</span>
                </div>
            )
        },
        {
            key: "status",
            name: "Status",
            fieldName: "status",
            minWidth: 120,
            maxWidth: 120,
            isResizable: false,
            onRender: (item: RecordType) => {
                return (
                    <div style={{ width: '120px', textAlign: 'center' }}>
                        <span style={{ cursor: 'default', paddingRight: '80px' }}>
                            {item.status}
                        </span>
                    </div>
                );
            }
        },
        {
            key: "category",
            name: "Category",
            fieldName: "category",
            minWidth: 100,
            maxWidth: 100,
            isResizable: false,
            onRender: (item: RecordType) => (
                <div style={{ width: '100px' }}>
                    <span style={{ cursor: 'default' }}>
                        {item.category}
                    </span>
                </div>
            )
        }
    ];

    // ---------------- UI ----------------
    return (
        <div style={{
            // height: "100vh",
            width: "100%",
            overflow: "hidden",
            display: "flex",
            flexDirection: "column",
            border: "1px solid #d1d1d1",
            borderRadius: "8px",
            backgroundColor: "#ffffff",
            boxShadow: "0 1px 3px rgba(0, 0, 0, 0.1)"
        }}>

            <CommandBar 
                items={commandBarItems} 
                styles={{
                    root: {
                        borderBottom: "1px solid #d1d1d1",
                        height: "40px",
                        minHeight: "40px",
                        padding: "0 8px"
                    }
                }}
            />

            <div style={{
                flex: 1,
                overflow: "auto",
                background: "white"
            }}>

                <DetailsList
                    items={groupedData}
                    columns={columns}
                    groups={groups}
                    selection={selectionRef.current}
                    selectionMode={SelectionMode.multiple}
                    checkboxVisibility={CheckboxVisibility.always}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    constrainMode={ConstrainMode.horizontalConstrained}
                    useReducedRowRenderer={true}

                    styles={{
                        root: {
                            overflowX: "auto"
                        }
                    }}

                    onRenderRow={(props, defaultRender) => {
                        if (!props) return null;

                        const record = props.item as RecordType;
                        const className = record.quantity < 10
                            ? "fluent-grid-row-red"    // Light red for qty < 10
                            : "fluent-grid-row-green"; // Light green for qty >= 10

                        return (
                            <div 
                                className={className}
                                onDoubleClick={() => {
                                    console.log("Double-clicked on row:", record.id);
                                    setEditingRows(prev => new Set([...prev, record.id]));
                                }}
                                style={{ cursor: 'pointer' }}
                            >
                                {defaultRender ? defaultRender(props) : null}
                            </div>
                        );
                    }}

                    groupProps={{
                        onRenderHeader: (props?: IGroupHeaderProps) => {
                            if (!props) return null;

                            const totals = props.group?.data;

                            return (
                                <div style={{ position: 'relative' }}>
                                    <GroupHeader
                                        {...props}
                                        styles={{
                                            root: {
                                                backgroundColor: '#f3f2f1',
                                                color: '#323130',
                                                padding: '0',
                                                margin: '0',
                                                fontWeight: 600,
                                                display: 'flex',
                                                alignItems: 'center',
                                                height: '32px'
                                            }
                                        }}
                                    />

                                    {/* Overlay totals in the group header aligned with columns */}
                                    <div style={{
                                        position: 'absolute',
                                        top: '0',
                                        left: '0',
                                        right: '0',
                                        height: '100%',
                                        display: 'flex',
                                        alignItems: 'center',
                                        pointerEvents: 'none'
                                    }}>
                                        {/* Checkbox + expand icon space - match exact spacing */}
                                        <div style={{ width: '48px' }}></div>

                                        {/* Product column - skip */}
                                        <div style={{ width: '270px' }}></div>

                                        {/* Quantity column total */}
                                        <div style={{
                                            width: '80px',
                                            textAlign: 'right',
                                            color: '#323130',
                                            fontSize: '12px',
                                            fontWeight: 600,
                                            paddingRight: '16px',
                                            paddingLeft: '8px'
                                        }}>
                                            {totals?.totalQty}
                                        </div>

                                        {/* Amount column total */}
                                        <div style={{
                                            width: '113px',
                                            textAlign: 'right',
                                            color: '#323130',
                                            fontSize: '12px',
                                            fontWeight: 600,
                                            paddingRight: '16px',
                                            paddingLeft: '8px'
                                        }}>
                                            ${totals?.totalAmt?.toFixed(2)}
                                        </div>
                                    </div>
                                </div>
                            );
                        }
                    }}
                />

            </div>
        </div>
    );
};