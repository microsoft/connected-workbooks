# Connected Workbooks
[![Build Status](https://obilshield.visualstudio.com/ConnectedWorkbooks/_apis/build/status/microsoft.connected-workbooks?branchName=main)](https://obilshield.visualstudio.com/ConnectedWorkbooks/_build/latest?definitionId=14&branchName=main)

## Usage

Use `connected-workbooks` to generate Excel workbooks with Power Query in them, mainly targeting "Export to Excel" features you have in your application.

### Basic Usage

#### Using a predefined template:
`connected-workbooks` comes with a default empty workbook template built in, that loads a query named "Query1" to a Query Table on the grid.

```typescript
let workbookManager = new WorkbookManager();

let blob = await workbookManager.generateSingleQueryWorkbook({
  query: {
    queryMashup,
    refreshOnOpen
  }
});

download(blob, filename);
```

<details>
<summary>download function example</summary>

Here's a typical way to download a file in browser with Typescript.

```typescript
function download(file: Blob, filename: string) {
  if (window.navigator.msSaveOrOpenBlob)
    // IE10+
    window.navigator.msSaveOrOpenBlob(file, filename);
  else {
    // Others
    var a = document.createElement("a"),
      url = URL.createObjectURL(file);
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }, 0);
  }
}
```
</details>

</br>

### Advanced Usage

#### Bring your own template:

You can use the library with your own template, for that (currently) there should be a single query named **Query1** loaded to a **Query Table**, **Pivot Table**, or **Pivot Chart** in the template workbook, and pass it as a **File** to the templateWorkbook parameter of the API:

```typescript
let blob = await workbookManager.generateSingleQueryWorkbook(
  query: {
    queryMashup,
    refreshOnOpen
  },
  templateFile
);
```

<details>
<summary>React example: file input</summary>
  
A common way to get the template workbook with React via user interaction:

```typescript
  const [templateFile, setTemplateFile] = useState<File | null>(null);

...

    <input
      onChange={(e) => {
          if (e?.target?.files?.item(0) == null) return;
          setTemplateFile(e!.target!.files!.item(0));
      }}
      type="file"
      id="file"
      accept=".xlsx"
      style={{ display: "none" }}
    />

```
</details>

Though expecation is that you have your product template baked and tested in your own product code, and not have the user provide it.

#### Customize docProps:
You can provide your own docProps (document properties) for the generated workbook, whether you are using the default or a custom template.

```typescript
let blob = await workbookManager.generateSingleQueryWorkbook(
  query: {
    queryMashup,
    refreshOnOpen
  },
  docProps: {
    title,
    subject,
    keywords,
    createdBy,
    description,
    lastModifiedBy,
    category,
    revision,
  }
);
```

### API
`connected-workbook` exposes __WorkbookManager__ class, which has (for now) a single method.


#### async `generateSingleQueryWorkbook`: `Promise<Blob>`

|param   | type   | required   | description   |
|---      |---    |---          |---            |
|query   | [QueryInfo](#queryinfo)   | __required__  | Power Query mashup  | 
| templateFile  | File   | optional   | Custom Excel workbook  | 
| docProps  | [DocProps](#docprops)   | optional  | Custom workbook properties |

</br>

### Types

### QueryInfo
| param | type | required | description |
|---|---|---|---|
| queryMashup | string | __required__ | mashup string
| refreshOnOpen | boolean | __required__ | Whether to refresh the data on opening workbook or not

### DocProps
| param | type | required | 
|---|---|---|
| title | string | optional 
| subject | string | optional 
| keywords | string | optional 
| createdBy | string | optional 
| description | string | optional 
| lastModifiedBy | string | optional 
| category | string | optional 
| revision | number | optional 

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
