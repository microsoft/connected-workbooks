<div align="center">

# Open In Excel

[![License](https://img.shields.io/github/license/microsoft/connected-workbooks)](https://github.com/microsoft/connected-workbooks/blob/master/LICENSE)
[![npm version](https://img.shields.io/npm/v/@microsoft/connected-workbooks)](https://www.npmjs.com/package/@microsoft/connected-workbooks)
[![Build Status](https://img.shields.io/github/workflow/status/microsoft/connected-workbooks/CI)](https://github.com/microsoft/connected-workbooks/actions)

**Open your data directly in Excel for the Web with zero installation** - A JavaScript library that converts web tables and data into interactive Excel workbooks with Power Query integration and custom branded templates

<div align="center">
<img src="./assets/template example.gif" alt="alt text" width="500" height="300">
</div>
</div>

---

## ✨ What Makes This Special?

🎯 **Interactive Workbooks, Not Static Files** - Generate fully functional Excel workbooks with tables instead of basic CSV exports that lose all structure and functionality.

🌐 **Zero-Install Excel Experience** - Launch workbooks directly in Excel for the Web through any browser without requiring Excel desktop installation, making your data accessible to any user anywhere.

🎨 **Corporate Branding & Custom Dashboards** - Inject your data into pre-built Excel templates containing your company branding, PivotTables, charts, and business logic while preserving all formatting and calculations.

🔄 **Live Data Connections with Power Query** - Create workbooks that automatically refresh from your web APIs, databases, or data sources using Microsoft's Power Query technology, eliminating manual data updates.

---

## 🌟 Key Features

| Feature | Description |
|---------|-------------|
| **📊 Smart Tables** | Export HTML tables or raw data to Excel tables  |
| **📱 Web Integration** | Open workbooks directly in Excel for the Web |
| **🎨 Custom Templates** | Bring your own branded Excel templates with PivotTables and charts |
| **🔗 Live Connections** | Create workbooks that refresh data on-demand using Power Query |
| **⚙️ Advanced Config** | Control document properties, sheet names, and table settings |

---
[Learn about Power Query here](https://powerquery.microsoft.com/en-us/)

## Where is this library used? here are some examples:

|<img src="https://github.com/microsoft/connected-workbooks/assets/7674478/b7a0c989-7ba4-4da8-851e-04650d8b600e" alt="Kusto" width="32"/>| <img src="https://github.com/microsoft/connected-workbooks/assets/7674478/76d22d23-5f2b-465f-992d-f1c71396904c" alt="LogAnalytics" width="32"/>	| <img src="https://github.com/microsoft/connected-workbooks/assets/7674478/436b4f53-bf25-4c45-aae5-55ee1b1feafc" alt="Datamart" width="32"/>	| <img src="https://github.com/microsoft/connected-workbooks/assets/7674478/3965f684-b461-42fe-9c62-e3059c0286eb" alt="VivaSales" width="32"/>	|
|---------------------------------	|-------------------	|--------------	|----------------	|
| **Azure Data Explorer** 	| **Log Analytics** 	| **Datamart** 	| **Viva Sales** 	|


---

## 🚀 Quick Start

### Installation

```bash
npm install @microsoft/connected-workbooks
```


---

## 💡 Usage Examples

### 1. 📋 Export HTML Table (Zero Configuration)

```typescript
import { workbookManager } from '@microsoft/connected-workbooks';

const blob = await workbookManager.generateTableWorkbookFromHtml(
  document.querySelector('table') as HTMLTableElement
);    
workbookManager.openInExcelWeb(blob, "MyTable.xlsx", true /*allowTyping*/);
```

### 2. 📊 Export Raw Data with Smart Formatting

```typescript
import { workbookManager } from '@microsoft/connected-workbooks';

const grid = {
  config: { promoteHeaders: true, adjustColumnNames: true },
  data: [
    ["Product", "Price", "InStock", "Category", "Date"],
    ["Widget A", 19.99, true, "Electronics", "10/26/2024"],
    ["Gizmo B", 9.99, true, "Accessories", "10/26/2024"],
    ["Bubala", 14.99, false, "Accessories", "10/22/2023"],
    ["Thingamajig C", 50, false, "Tools", "5/12/2023"],
    ["Doohickey D", 50.01, true, "Home", "8/12/2023"]
  ]
};

const blob = await workbookManager.generateTableWorkbookFromGrid(grid);    
workbookManager.openInExcelWeb(blob, "MyData.xlsx", true);
```

<div align="center">
<img width="350" alt="Generated Excel Table" src="https://github.com/microsoft/connected-workbooks/assets/7674478/b91e5d69-8444-4a19-a4b0-3fd721e5576f">
</div>

### 3. 📄 Custom Document Properties

```typescript
const blob = await workbookManager.generateTableWorkbookFromHtml(
  document.querySelector('table') as HTMLTableElement, 
  {
    docProps: { 
      createdBy: 'John Doe',
      lastModifiedBy: 'Jane Doe',
      description: 'Sales Report Q4 2024',
      title: 'Quarterly Sales Data'
    }
  }
);
     
workbookManager.downloadWorkbook(blob, "SalesReport.xlsx");
```

<div align="center">
<img width="400" alt="Document Properties" src="https://github.com/microsoft/connected-workbooks/assets/7674478/c267c9eb-6367-419d-832d-5a835c7683f9">
</div>

### 4. 🔄 Power Query Connected Workbook

Create workbooks that refresh data automatically:

```typescript
import { workbookManager } from '@microsoft/connected-workbooks';

const blob = await workbookManager.generateSingleQueryWorkbook({
  queryMashup: `let 
    Source = {1..10} 
  in 
    Source`,
  refreshOnOpen: true
});

workbookManager.openInExcelWeb(blob, "MyData.xlsx", true);
```

<div align="center">
<img width="100" alt="Power Query Workbook" src="https://github.com/microsoft/connected-workbooks/assets/7674478/57bd986c-6309-4963-8d86-911ccf496c3f">
</div>

### 5. 🎨 Advanced: Custom Branded Templates

Bring your own Excel template with pre-built dashboards:

📁 Template Loading Methods

```typescript
// Loading files
// Method 1: File input from user
const templateFile = document.querySelector('#template-upload').files[0];

// Method 2: Fetch from server
const templateResponse = await fetch('***.xlsx');
const templateFile = await templateResponse.blob();

// Method 3: From drag & drop
function handleTemplateDrop(event: DragEvent) {
  const templateFile = event.dataTransfer.files[0];
}

// Choose the data
const grid = {
  config: { promoteHeaders: true, adjustColumnNames: true },
  data: [
    ["Product", "Price", "InStock", "Category", "Date"],
    ["Widget A", 19.99, true, "Electronics", "10/26/2024"],
    ["Gizmo B", 9.99, true, "Accessories", "10/26/2024"],
    ["Bubala", 14.99, false, "Accessories", "10/22/2023"],
    ["Thingamajig C", 50, false, "Tools", "5/12/2023"],
    ["Doohickey D", 50.01, true, "Home", "8/12/2023"]
  ]
};

// Generate Workbook
const blob = await workbookManager.generateTableWorkbookFromGrid(
  grid,
  undefined, // no initial data grid
  { 
    templateFile: myCustomTemplate,
    TempleteSettings: {
      sheetName: "Dashboard", // Optional 
      tableName: "DataSource" // Optional 
    }
  }
);

// Download file
workbookManager.downloadWorkbook(blob, "BrandedReport.xlsx");

//Open in excel web
workbookManager.openInExcelWeb(blob, "MyData.xlsx", true);
```

<div align="center">
<img width="500" alt="Custom Branded Excel Template with Power Query Integration" src="https://github.com/microsoft/connected-workbooks/assets/7674478/e5377946-4348-4229-9b88-1910ff7ee025">
</div>

> 💡 **Template Requirements**: Include a query named **"Query1"** connected to a **Table**, **PivotTable**, or **PivotChart**.

## 📚 Complete API Reference

### Core Functions

#### 🔗 `generateSingleQueryWorkbook()`
Create Power Query connected workbooks with live data refresh capabilities.

```typescript
async function generateSingleQueryWorkbook(
  query: QueryInfo, 
  grid?: Grid, 
  fileConfigs?: FileConfigs
): Promise<Blob>
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `query` | [`QueryInfo`](#queryinfo) | ✅ **Required** | Power Query configuration |
| `grid` | [`Grid`](#grid) |  Optional | Pre-populate with data |
| `fileConfigs` | [`FileConfigs`](#fileconfigs) |  Optional | Customization options |

#### 📋 `generateTableWorkbookFromHtml()`
Convert HTML tables to Excel workbooks instantly.

```typescript
async function generateTableWorkbookFromHtml(
  htmlTable: HTMLTableElement, 
  fileConfigs?: FileConfigs
): Promise<Blob>
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `htmlTable` | `HTMLTableElement` | ✅ **Required** | Source HTML table |
| `fileConfigs` | [`FileConfigs`](#fileconfigs) |  Optional | Customization options |

#### 📊 `generateTableWorkbookFromGrid()`
Transform raw data arrays into formatted Excel tables.

```typescript
async function generateTableWorkbookFromGrid(
  grid: Grid, 
  fileConfigs?: FileConfigs
): Promise<Blob>
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `grid` | [`Grid`](#grid) | ✅ **Required** | Data and configuration |
| `fileConfigs` | [`FileConfigs`](#fileconfigs) |  Optional | Customization options |

#### 🌐 `openInExcelWeb()`
Open workbooks directly in Excel for the Web.

```typescript
async function openInExcelWeb(
  blob: Blob, 
  filename?: string, 
  allowTyping?: boolean
): Promise<void>
```

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `blob` | `Blob` | ✅ **Required** | Generated workbook |
| `filename` | `string` |  Optional | Custom filename |
| `allowTyping` | `boolean` |  Optional | Enable editing (default: false) |

#### 💾 `downloadWorkbook()`
Trigger browser download of the workbook.

```typescript
function downloadWorkbook(file: Blob, filename: string): void
```

#### 🔗 `getExcelForWebWorkbookUrl()` 
Get the Excel for Web URL without opening (useful for custom integrations).

```typescript
async function getExcelForWebWorkbookUrl(
  file: Blob, 
  filename?: string, 
  allowTyping?: boolean
): Promise<string>
```

---

## 🔧 Type Definitions

### QueryInfo
Power Query configuration for connected workbooks.

```typescript
interface QueryInfo {
  queryMashup: string;        // Power Query M language code
  refreshOnOpen: boolean;     // Auto-refresh when opened
  queryName?: string;         // Query identifier (default: "Query1")
}
```

### Grid
Data structure for tabular information.

```typescript
interface Grid {
  data: (string | number | boolean)[][];  // Raw data rows
  config?: GridConfig;                    // Processing options
}

interface GridConfig {
  promoteHeaders?: boolean;     // Use first row as headers
  adjustColumnNames?: boolean;  // Fix duplicate/invalid names
}
```

### FileConfigs
Advanced customization options.

```typescript
interface FileConfigs {
  templateFile?: File;              // Custom Excel template
  docProps?: DocProps;              // Document metadata
  hostName?: string;                // Creator application name
  TempleteSettings?: TempleteSettings; // Template-specific settings
}

interface TempleteSettings {
  tableName?: string;    // Target table name in template
  sheetName?: string;    // Target worksheet name
}
```

### DocProps
Document metadata and properties.

```typescript
interface DocProps {
  title?: string;           // Document title
  subject?: string;         // Document subject
  keywords?: string;        // Search keywords
  createdBy?: string;       // Author name
  description?: string;     // Document description
  lastModifiedBy?: string;  // Last editor
  category?: string;        // Document category
  revision?: string;        // Version number
}
```

---


## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

### Getting Started
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

### Development Setup
```bash
git clone https://github.com/microsoft/connected-workbooks.git
cd connected-workbooks
npm install
npm run build
npm test
```
---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🔗 Related Resources

- [📖 Power Query Documentation](https://powerquery.microsoft.com/en-us/)
- [🏢 Excel for Developers](https://docs.microsoft.com/en-us/office/dev/excel/)
- [🔧 Microsoft Graph Excel APIs](https://docs.microsoft.com/en-us/graph/api/resources/excel)

---

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft
trademarks or logos is subject to and must follow
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.

---

## Keywords

Power Query, Excel, Office, Workbook, Refresh, Table, xlsx, export, CSV, data export, HTML table, web to Excel, JavaScript Excel, TypeScript Excel, Excel template, PivotTable, connected data, live data, data refresh, Excel for Web, browser Excel, spreadsheet, data visualization, Microsoft Office, Office 365, Excel API, workbook generation, table export, grid export, Excel automation, data processing, business intelligence
