<div align="center">

# Connected Workbooks (TypeScript)

[![TypeScript](https://img.shields.io/badge/TypeScript-007ACC?style=flat&logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![npm version](https://img.shields.io/npm/v/@microsoft/connected-workbooks)](https://www.npmjs.com/package/@microsoft/connected-workbooks)
[![Build Status](https://img.shields.io/github/actions/workflow/status/microsoft/connected-workbooks/azure-pipelines.yml?branch=main)](https://github.com/microsoft/connected-workbooks/actions)

</div>

> Need the high-level overview, feature tour, or repo structure? See the root [`README.md`](../README.md). This file covers the TypeScript package specifically.

---

## Installation

```bash
npm install @microsoft/connected-workbooks
```

The library targets evergreen browsers. No native modules, build steps, or Office add-ins are required.

---

## Usage Examples

### HTML Table Export

```ts
import { workbookManager } from '@microsoft/connected-workbooks';

const table = document.querySelector('table') as HTMLTableElement;
const blob = await workbookManager.generateTableWorkbookFromHtml(table);
await workbookManager.openInExcelWeb(blob, 'QuickExport.xlsx', true);
```

### Grid Export With Smart Headers

```ts
const grid = {
  config: {
    promoteHeaders: true,
    adjustColumnNames: true
  },
  data: [
    ['Product', 'Revenue', 'InStock', 'Category'],
    ['Surface Laptop', 1299.99, true, 'Hardware'],
    ['Office 365', 99.99, true, 'Software']
  ]
};

const blob = await workbookManager.generateTableWorkbookFromGrid(grid);
await workbookManager.openInExcelWeb(blob, 'SalesReport.xlsx', true);
```

### Inject Data Into Templates

1. Design an `.xlsx` template that includes a table (for example `QuarterlyData`).
2. Upload or fetch the template in the browser.
3. Supply the file plus metadata:

```ts
const templateResponse = await fetch('/assets/templates/sales-dashboard.xlsx');
const templateFile = await templateResponse.blob();

const blob = await workbookManager.generateTableWorkbookFromGrid(
  quarterlyData,
  undefined,
  {
    templateFile,
    TempleteSettings: {
      sheetName: 'Dashboard',
      tableName: 'QuarterlyData'
    }
  }
);

await workbookManager.openInExcelWeb(blob, 'ExecutiveDashboard.xlsx', true);
```

### Power Query Workbooks

```ts
const blob = await workbookManager.generateSingleQueryWorkbook({
  queryMashup: `let Source = Json.Document(Web.Contents('https://contoso/api/orders')) in Source`,
  refreshOnOpen: true
});

await workbookManager.openInExcelWeb(blob, 'Orders.xlsx', true);
```

### Document Properties & Downloads

```ts
const blob = await workbookManager.generateTableWorkbookFromHtml(table, {
  docProps: {
    createdBy: 'Contoso Portal',
    description: 'Q4 pipeline export',
    title: 'Executive Dashboard'
  }
});

workbookManager.downloadWorkbook(blob, 'Pipeline.xlsx');
```

---

## API Surface

### `generateSingleQueryWorkbook()`

```ts
async function generateSingleQueryWorkbook(
  query: QueryInfo,
  grid?: Grid,
  fileConfigs?: FileConfigs
): Promise<Blob>
```

- `query`: Power Query definition (M script, refresh flag, query name).
- `grid`: Optional seed data.
- `fileConfigs`: Template, metadata, or host options.

### `generateTableWorkbookFromHtml()`

```ts
async function generateTableWorkbookFromHtml(
  htmlTable: HTMLTableElement,
  fileConfigs?: FileConfigs
): Promise<Blob>
```

### `generateTableWorkbookFromGrid()`

```ts
async function generateTableWorkbookFromGrid(
  grid: Grid,
  fileConfigs?: FileConfigs
): Promise<Blob>
```

### `openInExcelWeb()` and `getExcelForWebWorkbookUrl()`

Launch Excel for the Web immediately or just capture the URL for custom navigation flows.

### `downloadWorkbook()`

Trigger a regular browser download of the generated blob.

---

## Type Definitions

```ts
interface QueryInfo {
  queryMashup: string;
  refreshOnOpen: boolean;
  queryName?: string; // default: "Query1"
}

interface Grid {
  data: (string | number | boolean | null)[][];
  config?: {
    promoteHeaders?: boolean;
    adjustColumnNames?: boolean;
  };
}

interface FileConfigs {
  templateFile?: File | Blob | Buffer;
  docProps?: DocProps;
  hostName?: string;
  TempleteSettings?: {
    tableName?: string;
    sheetName?: string;
  };
}

interface DocProps {
  title?: string;
  subject?: string;
  keywords?: string;
  createdBy?: string;
  description?: string;
  lastModifiedBy?: string;
  category?: string;
  revision?: string;
}
```

---

## Development

```bash
cd typescript
npm install
npm run build
npm test
```

Use `npm run validate:implementations` to compare the TypeScript and .NET output when making cross-language changes.

---

## Contributing

Follow the guidance in the root [`README.md`](../README.md#contributing). Pull requests should include unit tests (`npm test`) and adhere to the repo ESLint/Prettier settings before submission.
