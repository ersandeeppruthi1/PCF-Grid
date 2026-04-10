# PCF FluentUI Data Grid Control - AI Coding Guide

## Project Overview
This is a **PowerApps Component Framework (PCF)** control that creates a sophisticated data grid using **FluentUI React components** for Microsoft Dynamics 365/Power Platform integration.

## Architecture & Key Patterns

### Core Structure
- **`FluentGridControl.pcfproj`**: MSBuild project file for PCF compilation
- **`FluentGridControl/index.ts`**: PCF lifecycle implementation & React bridge
- **`FluentGridControl/FluentGrid.tsx`**: Main React component with grid functionality
- **`FluentGridControl/ControlManifest.Input.xml`**: PCF metadata & capabilities
- **`generated/ManifestTypes.d.ts`**: Auto-generated TypeScript interfaces

### Critical PCF-React Bridge Pattern
```typescript
// ✅ ESSENTIAL: Use React wrapper to prevent re-render loops
class GridWrapper extends React.Component<{ data: RecordType[] }> {
    render() {
        return React.createElement(FluentGrid, { data: this.props.data });
    }
}

// ✅ ESSENTIAL: Hash-based change detection prevents infinite re-renders
const newHash = JSON.stringify(records);
if (this.lastDataHash === newHash) return; // STOP LOOP
this.lastDataHash = newHash;
```

### Data Binding & Context Integration
- **Dataset Mapping**: `context.parameters.dataset` → transformed `RecordType[]`
- **Field Mapping**: Handles both direct values and lookup references:
  ```typescript
  const productRef = record.getValue("productname") as ComponentFramework.LookupValue[] | null;
  const productName = productRef?.[0]?.name ?? String(record.getValue("productname") ?? "N/A");
  ```

## Development Workflows

### Building & Testing
```bash
npm run build          # Compile PCF control
npm run start         # Start test harness
npm run start:watch   # Watch mode for development
pac pcf push          # Deploy to environment
```

### PCF-Specific Debugging
- **Test Harness**: Uses `npm start` for local React development
- **Dynamics 365 Testing**: Deploy with `pac pcf push` to test real dataset binding
- **Console Logging**: Key debug point in `updateView()` method

## Project-Specific Conventions

### State Management Pattern
```typescript
// ✅ Separate editing vs dirty tracking
const [editingRows, setEditingRows] = React.useState<Set<string>>(new Set());
const [dirtyRecords, setDirtyRecords] = React.useState<Set<string>>(new Set());

// ✅ Field updates always mark as dirty
const updateField = (id: string, field: keyof RecordType, value: string | number) => {
    setDirtyRecords(prev => new Set([...prev, id]));
    setData(prev => prev.map(r => r.id === id ? { ...r, [field]: value, isDirty: true } : r));
};
```

### Dynamics 365 Web API Integration
```typescript
// ✅ Standard D365 Web API pattern
const apiUrl = `${baseUrl}/api/data/v9.2/${entityName}(${record.id})`;
const response = await fetch(apiUrl, {
    method: 'PATCH',
    headers: {
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'If-Match': '*' // Optimistic concurrency
    }
});
```

### FluentUI Styling Overrides
```typescript
// ✅ Dynamic style injection for row coloring
React.useEffect(() => {
    const styleElement = document.createElement('style');
    styleElement.innerHTML = `
        .fluent-grid-row-red { background-color: #ffcccb !important; }
        .fluent-grid-row-green { background-color: #90ee90 !important; }
    `;
    document.head.appendChild(styleElement);
    return () => styleElement.remove();
}, []);
```

## Key Integration Points

### PCF Lifecycle Hooks
- **`init()`**: Initialize container reference
- **`updateView()`**: Handle dataset changes & prevent re-render loops
- **`destroy()`**: Clean up React components

### Platform Libraries
- **React 16.14.0**: Declared in manifest, accessed via global
- **FluentUI 9.46.2**: Platform-provided, imported from `@fluentui/react`
- **WebAPI & Utility**: Required PCF features for D365 integration

### Data Flow Architecture
1. **Dynamics Dataset** → `context.parameters.dataset`
2. **PCF Transform** → `RecordType[]` with lookup resolution
3. **React State** → Grid rendering with edit capabilities
4. **Web API Sync** → Batch updates back to Dynamics

## Common Pitfalls & Solutions

- **Re-render Loops**: Always use hash-based change detection in `updateView()`
- **Lookup Fields**: Handle both `LookupValue[]` and `LookupValue` types
- **Style Conflicts**: Use `!important` and dynamic style injection for FluentUI overrides
- **Async Operations**: Batch Web API calls with `Promise.allSettled()` for reliability

## File Structure Conventions
```
FluentGridControl/
├── FluentGridControl/          # Control source code
│   ├── index.ts               # PCF lifecycle & React bridge
│   ├── FluentGrid.tsx         # Main React component
│   ├── ControlManifest.Input.xml  # PCF configuration
│   └── generated/             # Auto-generated types
├── out/controls/              # Build output
└── obj/                       # MSBuild artifacts
```
