# Connected Workbooks
[![Build Status](https://obilshield.visualstudio.com/ConnectedWorkbooks/_apis/build/status/microsoft.connected-workbooks?branchName=main)](https://obilshield.visualstudio.com/ConnectedWorkbooks/_build/latest?definitionId=14&branchName=main)
[![License](https://img.shields.io/github/license/microsoft/connected-workbooks)](https://github.com/microsoft/connected-workbooks/blob/master/LICENSE)
[![Snyk Vulnerabilities](https://img.shields.io/snyk/vulnerabilities/github/microsoft/connected-workbooks)](https://snyk.io/test/github/microsoft/connected-workbooks)

A pure JS library, Microsoft backed, that provides xlsx workbook generation capabilities, allowing for:
1. Fundemental **"Export to Excel"** capabilities for tabular data (landing in a table in Excel).
2. Advanced capabilities of **"Export a Power Query connected workbook"**:
    - Can refresh your data on open and/or on demand.
    - Allows for initial data population.
    - Supports more advanced scenarios where you provide branded/custom workbooks, and load your data into PivotTables or PivotCharts.

Connected Workbooks allows you to avoid "data dumps" in CSV form, providing a richer experience with Tables and/or connected Queries for when your business application supports it.

[Learn about Power Query here](https://powerquery.microsoft.com/en-us/)
[![License](https://img.shields.io/github/license/microsoft/connected-workbooks)](https://github.com/microsoft/connected-workbooks/blob/master/LICENSE)
[![Snyk Vulnerabilities](https://img.shields.io/snyk/vulnerabilities/github/microsoft/connected-workbooks)](https://snyk.io/test/github/microsoft/connected-workbooks)

A pure JS library, Microsoft backed, that provides xlsx workbook generation capabilities, allowing for:
1. Fundemental **"Export to Excel"** capabilities for tabular data (landing in a table in Excel).
2. Advanced capabilities of **"Export a Power Query connected workbook"**:
    - Can refresh your data on open and/or on demand.
    - Allows for initial data population.
    - Supports more advanced scenarios where you provide branded/custom workbooks, and load your data into PivotTables or PivotCharts.

Connected Workbooks allows you to avoid "data dumps" in CSV form, providing a richer experience with Tables and/or connected Queries for when your business application supports it.

[Learn about Power Query here](https://powerquery.microsoft.com/en-us/)

## Where is this library used? here are some examples:
## Where is this library used? here are some examples:

|<img src="https://github.com/microsoft/connected-workbooks/assets/7674478/b7a0c989-7ba4-4da8-851e-04650d8b600e" alt="Kusto" width="32"/>| <img src="https://github.com/microsoft/connected-workbooks/assets/7674478/76d22d23-5f2b-465f-992d-f1c71396904c" alt="LogAnalytics" width="32"/>	| <img src="https://github.com/microsoft/connected-workbooks/assets/7674478/436b4f53-bf25-4c45-aae5-55ee1b1feafc" alt="Datamart" width="32"/>	| <img src="https://github.com/microsoft/connected-workbooks/assets/7674478/3965f684-b461-42fe-9c62-e3059c0286eb" alt="VivaSales" width="32"/>	|
|---------------------------------	|-------------------	|--------------	|----------------	|
| **Azure Data Explorer** 	| **Log Analytics** 	| **Datamart** 	| **Viva Sales** 	|

## How do I use it? here are some examples:

### 1. Export a table directly from an Html page:
```typescript
import { workbookManager } from '@microsoft/connected-workbooks';
### 1. Export a table directly from an Html page:
```typescript
import { workbookManager } from '@microsoft/connected-workbooks';

const blob = await workbookManager.generateTableWorkbookFromHtml(document.querySelector('table') as HTMLTableElement);    
workbookManager.downloadWorkbook(blob, "MyTable.xlsx");
```
### 2. Export a table from raw data:
const blob = await workbookManager.generateTableWorkbookFromHtml(document.querySelector('table') as HTMLTableElement);    
workbookManager.downloadWorkbook(blob, "MyTable.xlsx");
```
### 2. Export a table from raw data:
```typescript
import { workbookManager } from '@microsoft/connected-workbooks';

const grid = {
  "promoteHeaders": false,
  "gridData": [
      ["Product", "Price", "InStock", "Category", "Date"],
      ["Widget A", 19.99, true, "Electronics", "10/26/2024"],
      ["Gizmo B", 9.99, true, "Accessories", "10/26/2024"],
      ["Bubala", 14.99, false, "Accessories", "10/22/2023"],
      ["Thingamajig C", 50, false, "Tools", "5/12/2023"],
      ["Doohickey D", 50.01, true, "Home", "8/12/2023"]
  ]
};
const blob = await workbookManager.generateTableWorkbookFromGrid(grid);    
workbookManager.downloadWorkbook(blob, "MyTable.xlsx");
```
<img width="281" alt="image" src="https://github.com/microsoft/connected-workbooks/assets/7674478/b91e5d69-8444-4a19-a4b0-3fd721e5576f">

### 3. Control Document Properties:
```typescript
    const blob = await workbookManager.generateTableWorkbookFromHtml(
      document.querySelector('table') as HTMLTableElement, 
      {createdBy: 'John Doe', lastModifiedBy: 'Jane Doe', description: 'This is a sample table'});
    
      workbookManager.downloadWorkbook(blob, "MyTable.xlsx");
```
![image](https://github.com/microsoft/connected-workbooks/assets/7674478/c267c9eb-6367-419d-832d-5a835c7683f9)

### 4. Export a Power Query connected workbook:
```typescript
import { workbookManager } from '@microsoft/connected-workbooks';

const blob = await workbookManager.generateSingleQueryWorkbook({
  queryMashup: 'let \
                    Source = {1..10} \
                in \
                    Source',
  refreshOnOpen: true});
workbookManager.downloadWorkbook(blob, "MyConnectedWorkbook.xlsx");
});
import { workbookManager } from '@microsoft/connected-workbooks';

const blob = await workbookManager.generateSingleQueryWorkbook({
  queryMashup: 'let \
                    Source = {1..10} \
                in \
                    Source',
  refreshOnOpen: true});
workbookManager.downloadWorkbook(blob, "MyConnectedWorkbook.xlsx");
});
```
![image](https://github.com/microsoft/connected-workbooks/assets/7674478/57bd986c-6309-4963-8d86-911ccf496c3f)
(after refreshing on open)
### Advanced Usage - bring your own template:

You can use the library with your own workbook as a template!

```typescript
const blob = await workbookManager.generateSingleQueryWorkbook(
  { queryMashup: query, refreshOnOpen: true },
  undefined /* optional gridData */,
  templateFile);
workbookManager.downloadWorkbook(blob, "MyBrandedWorkbook.xlsx");
const blob = await workbookManager.generateSingleQueryWorkbook(
  { queryMashup: query, refreshOnOpen: true },
  undefined /* optional gridData */,
  templateFile);
workbookManager.downloadWorkbook(blob, "MyBrandedWorkbook.xlsx");
```
![image](https://github.com/microsoft/connected-workbooks/assets/7674478/e5377946-4348-4229-9b88-1910ff7ee025)

Template requirements:

Have a single query named **Query1** loaded to a **Query Table**, **Pivot Table**, or a **Pivot Chart**.


⭐ Recommendation - have your product template baked and tested in your own product code, instead of your user providing it.

⭐ For user templates - a common way to get the template workbook with React via user interaction:

```typescript
const [templateFile, setTemplateFile] = useState<File | null>(null);
const [templateFile, setTemplateFile] = useState<File | null>(null);
...
<input type="file" id="file" accept=".xlsx" style={{ display: "none" }} onChange={(e) => {
  if (e?.target?.files?.item(0) == null) return;
  setTemplateFile(e!.target!.files!.item(0));
}}/>
<input type="file" id="file" accept=".xlsx" style={{ display: "none" }} onChange={(e) => {
  if (e?.target?.files?.item(0) == null) return;
  setTemplateFile(e!.target!.files!.item(0));
}}/>
```

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

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft
trademarks or logos is subject to and must follow
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
