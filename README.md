# Connected Workbooks

## Using this project

Use this project in your to generate workbooks with Power Query in them, mainly targeting 'Export to Excel' features you have in your application.

### Basic Usage - using a predefined template:

The library comes with a workbook template built in, that loads a query named 'Query1' to a Query Table on the grid.

```typescript
let workbookManager = new WorkbookManager();
let blob = await workbookManager.generateSingleQueryWorkbook({
  queryMashup: query,
  refreshOnOpen: refreshOnOpen,
});

Download(blob, filename);
```

While a typical download method would be:

```typescript
function Download(file: Blob, filename: string) {
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

### Advanced Usage - bring your own template:

You can use the library with your own template, for that (currently) there should be a single query named **Query1** loaded to a **Query Table**, **Pivot Table**, or **Pivot Chart** in the template workbook, and pass it as a **File** to the templateWorkbook parameter of the API:

```typescript
let workbookManager = new WorkbookManager();
let blob = await workbookManager.generateSingleQueryWorkbook(
  {
    queryMashup: query,
    refreshOnOpen: refreshOnOpen,
  },
  templateFile
);

Download(blob, filename);
```

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

Though expecation is that you have your product template baked and tested in your own product code, and not have the user provide it.

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
