// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Microsoft.ConnectedWorkbooks.Models;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal sealed class WorkbookEditor
{
    private readonly ExcelArchive _archive;
    private readonly DocumentProperties? _documentProperties;

    public WorkbookEditor(ExcelArchive archive, DocumentProperties? documentProperties)
    {
        _archive = archive;
        _documentProperties = documentProperties;
    }

    public void UpdatePowerQueryDocument(string queryName, string mashupDocument)
    {
        var (path, base64) = PqUtilities.GetDataMashup(_archive);
        var nextBase64 = MashupDocumentParser.ReplaceSingleQuery(base64, queryName, mashupDocument);
        PqUtilities.SetDataMashup(_archive, path, nextBase64);
    }

    public string UpdateConnections(string queryName, bool refreshOnOpen)
    {
        var xml = _archive.ReadText(WorkbookConstants.ConnectionsXmlPath);
        var doc = XDocument.Parse(xml);
        var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
        var connection = doc.Root?.Element(ns + "connection") ?? throw new InvalidOperationException("Connections XML does not contain a connection element.");
        var dbPr = connection.Element(ns + "dbPr") ?? throw new InvalidOperationException("Connections XML is missing the dbPr element.");

        connection.SetAttributeValue("name", $"Query - {queryName}");
        connection.SetAttributeValue("description", $"Connection to the '{queryName}' query in the workbook.");
        dbPr.SetAttributeValue("connection", $"Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=\"{queryName}\";");
        dbPr.SetAttributeValue("command", $"SELECT * FROM [{queryName.Replace("]", "]]", StringComparison.Ordinal)}]");
        dbPr.SetAttributeValue("refreshOnLoad", refreshOnOpen ? "1" : "0");

        _archive.WriteText(WorkbookConstants.ConnectionsXmlPath, doc.ToString(SaveOptions.DisableFormatting));
        return connection.Attribute("id")?.Value ?? "1";
    }

    public int UpdateSharedStrings(string queryName)
    {
        var xml = _archive.ReadText(WorkbookConstants.SharedStringsXmlPath);
        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
        var textElements = doc.Descendants(ns + "t").ToList();
        var sharedStringIndex = textElements.Count;
        var existing = textElements.Select((element, index) => (element, index)).FirstOrDefault(tuple => tuple.element.Value == queryName);
        if (existing.element is not null)
        {
            sharedStringIndex = existing.index + 1;
        }
        else
        {
            var si = new XElement(ns + "si", new XElement(ns + "t", queryName));
            doc.Root!.Add(si);
            IncrementAttribute(doc.Root, "count");
            IncrementAttribute(doc.Root, "uniqueCount");
        }

        _archive.WriteText(WorkbookConstants.SharedStringsXmlPath, doc.ToString(SaveOptions.DisableFormatting));
        return sharedStringIndex;
    }

    public void UpdateWorksheet(int sharedStringIndex)
    {
        var xml = _archive.ReadText(WorkbookConstants.DefaultSheetPath);
        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
        var cellValue = doc.Descendants(ns + XmlNames.Elements.CellValue).FirstOrDefault();
        if (cellValue is null)
        {
            throw new InvalidOperationException("Worksheet XML did not contain a cell value node.");
        }

        cellValue.Value = sharedStringIndex.ToString(CultureInfo.InvariantCulture);
        _archive.WriteText(WorkbookConstants.DefaultSheetPath, doc.ToString(SaveOptions.DisableFormatting));
    }

    public void UpdateQueryTable(string connectionId, bool refreshOnOpen)
    {
        var xml = _archive.ReadText(WorkbookConstants.QueryTablePath);
        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
        doc.Root?.SetAttributeValue("connectionId", connectionId);
        doc.Root?.SetAttributeValue("refreshOnLoad", refreshOnOpen ? "1" : "0");
        _archive.WriteText(WorkbookConstants.QueryTablePath, doc.ToString(SaveOptions.DisableFormatting));
    }

    public void UpdateTableData(TableData tableData)
    {
        if (tableData.ColumnNames.Count == 0)
        {
            return;
        }

        UpdateSheetData(tableData);
        UpdateTableDefinition(tableData);
        UpdateWorkbookDefinedName(tableData);
        UpdateQueryTableColumns(tableData);
    }

    public void UpdateDocumentProperties()
    {
        var now = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
        var xml = _archive.ReadText(WorkbookConstants.DocPropsCoreXmlPath);
        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var cp = (XNamespace)"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        var dc = (XNamespace)"http://purl.org/dc/elements/1.1/";
        var dcterms = (XNamespace)"http://purl.org/dc/terms/";
        var xsi = (XNamespace)"http://www.w3.org/2001/XMLSchema-instance";

        SetElement(doc, cp + "coreProperties", dcterms + "created", now, xsi);
        SetElement(doc, cp + "coreProperties", dcterms + "modified", now, xsi);
        if (_documentProperties is not null)
        {
            SetElement(doc, cp + "coreProperties", dc + "title", _documentProperties.Title);
            SetElement(doc, cp + "coreProperties", dc + "subject", _documentProperties.Subject);
            SetElement(doc, cp + "coreProperties", dc + "creator", _documentProperties.CreatedBy);
            SetElement(doc, cp + "coreProperties", dc + "description", _documentProperties.Description);
            SetElement(doc, cp + "coreProperties", (XNamespace)"http://schemas.openxmlformats.org/package/2006/metadata/core-properties" + "keywords", _documentProperties.Keywords);
            SetElement(doc, cp + "coreProperties", cp + "lastModifiedBy", _documentProperties.LastModifiedBy);
            SetElement(doc, cp + "coreProperties", cp + "category", _documentProperties.Category);
            SetElement(doc, cp + "coreProperties", cp + "revision", _documentProperties.Revision);
        }

        _archive.WriteText(WorkbookConstants.DocPropsCoreXmlPath, doc.ToString(SaveOptions.DisableFormatting));
    }

    private void UpdateSheetData(TableData tableData)
    {
        var xml = _archive.ReadText(WorkbookConstants.DefaultSheetPath);
        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
        var x14ac = doc.Root?.GetNamespaceOfPrefix("x14ac") ?? XNamespace.None;
        var sheetData = doc.Descendants(ns + XmlNames.Elements.SheetData).FirstOrDefault();
        if (sheetData is null)
        {
            throw new InvalidOperationException("Worksheet XML is missing sheetData.");
        }

        sheetData.RemoveNodes();
        var startCell = "A1";
        var (startRow, startColumn) = CellReferenceHelper.GetStartPosition(startCell);
        var spans = $"{startColumn}:{startColumn + tableData.ColumnNames.Count - 1}";

        var headerRow = new XElement(ns + XmlNames.Elements.Row,
            new XAttribute(XmlNames.Attributes.Row, startRow),
            new XAttribute(XmlNames.Attributes.Spans, spans),
            x14ac == XNamespace.None ? null : new XAttribute(x14ac + "dyDescent", "0.3"));

        for (var columnIndex = 0; columnIndex < tableData.ColumnNames.Count; columnIndex++)
        {
            headerRow.Add(CreateCell(ns, startColumn + columnIndex, startRow, tableData.ColumnNames[columnIndex], isHeader: true));
        }

        sheetData.Add(headerRow);

        for (var rowIndex = 0; rowIndex < tableData.Rows.Count; rowIndex++)
        {
            var excelRow = startRow + rowIndex + 1;
            var row = new XElement(ns + XmlNames.Elements.Row,
                new XAttribute(XmlNames.Attributes.Row, excelRow),
                new XAttribute(XmlNames.Attributes.Spans, spans),
                x14ac == XNamespace.None ? null : new XAttribute(x14ac + "dyDescent", "0.3"));

            var rowValues = tableData.Rows[rowIndex];
            for (var columnIndex = 0; columnIndex < tableData.ColumnNames.Count; columnIndex++)
            {
                var value = columnIndex < rowValues.Count ? rowValues[columnIndex] : string.Empty;
                row.Add(CreateCell(ns, startColumn + columnIndex, excelRow, value, isHeader: false));
            }

            sheetData.Add(row);
        }

        var endReference = CellReferenceHelper.BuildReference((startRow, startColumn), tableData.ColumnNames.Count, tableData.Rows.Count + 1);
        doc.Descendants(ns + XmlNames.Elements.Dimension).FirstOrDefault()?.SetAttributeValue(XmlNames.Attributes.Reference, endReference);
        doc.Descendants(ns + XmlNames.Elements.Selection).FirstOrDefault()?.SetAttributeValue(XmlNames.Attributes.SqRef, endReference);

        _archive.WriteText(WorkbookConstants.DefaultSheetPath, doc.ToString(SaveOptions.DisableFormatting));
    }

    private void UpdateTableDefinition(TableData tableData)
    {
        var xml = _archive.ReadText(WorkbookConstants.DefaultTablePath);
        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
        var tableColumns = doc.Descendants(ns + XmlNames.Elements.TableColumns).FirstOrDefault();
        if (tableColumns is null)
        {
            throw new InvalidOperationException("Table definition is missing tableColumns.");
        }

        tableColumns.RemoveNodes();
        tableColumns.SetAttributeValue(XmlNames.Attributes.Count, tableData.ColumnNames.Count);
        for (var index = 0; index < tableData.ColumnNames.Count; index++)
        {
            var column = new XElement(ns + XmlNames.Elements.TableColumn);
            column.SetAttributeValue(XmlNames.Attributes.Id, index + 1);
            column.SetAttributeValue(XmlNames.Attributes.Name, tableData.ColumnNames[index]);
            tableColumns.Add(column);
        }

        var reference = $"A1:{CellReferenceHelper.ColumnNumberToName(tableData.ColumnNames.Count - 1)}{tableData.Rows.Count + 1}";
        doc.Root?.SetAttributeValue(XmlNames.Attributes.Reference, reference);
        doc.Descendants(ns + XmlNames.Elements.AutoFilter).FirstOrDefault()?.SetAttributeValue(XmlNames.Attributes.Reference, reference);

        _archive.WriteText(WorkbookConstants.DefaultTablePath, doc.ToString(SaveOptions.DisableFormatting));
    }

    private void UpdateWorkbookDefinedName(TableData tableData)
    {
        var xml = _archive.ReadText(WorkbookConstants.WorkbookXmlPath);
        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
        var definedName = doc.Descendants(ns + XmlNames.Elements.DefinedName).FirstOrDefault();
        if (definedName is null)
        {
            _archive.WriteText(WorkbookConstants.WorkbookXmlPath, doc.ToString(SaveOptions.DisableFormatting));
            return;
        }

        var reference = $"!$A$1:${CellReferenceHelper.ColumnNumberToName(tableData.ColumnNames.Count - 1)}${tableData.Rows.Count + 1}";
        definedName.Value = reference;
        _archive.WriteText(WorkbookConstants.WorkbookXmlPath, doc.ToString(SaveOptions.DisableFormatting));
    }

    private void UpdateQueryTableColumns(TableData tableData)
    {
        var xml = _archive.ReadText(WorkbookConstants.QueryTablePath);
        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
        var fields = doc.Descendants(ns + XmlNames.Elements.QueryTableFields).FirstOrDefault();
        if (fields is null)
        {
            throw new InvalidOperationException("Query table definition is missing queryTableFields.");
        }

        fields.RemoveNodes();
        for (var index = 0; index < tableData.ColumnNames.Count; index++)
        {
            var field = new XElement(ns + XmlNames.Elements.QueryTableField);
            field.SetAttributeValue(XmlNames.Attributes.Id, index + 1);
            field.SetAttributeValue(XmlNames.Attributes.Name, tableData.ColumnNames[index]);
            field.SetAttributeValue(XmlNames.Attributes.TableColumnId, index + 1);
            fields.Add(field);
        }

        fields.SetAttributeValue(XmlNames.Attributes.Count, tableData.ColumnNames.Count);
        doc.Descendants(ns + XmlNames.Elements.QueryTableRefresh).FirstOrDefault()?.SetAttributeValue(XmlNames.Attributes.NextId, tableData.ColumnNames.Count + 1);
        _archive.WriteText(WorkbookConstants.QueryTablePath, doc.ToString(SaveOptions.DisableFormatting));
    }

    private XElement CreateCell(XNamespace ns, int column, int row, string value, bool isHeader)
    {
        var reference = $"{CellReferenceHelper.ColumnNumberToName(column - 1)}{row}";
        var cell = new XElement(ns + XmlNames.Elements.Cell,
            new XAttribute("r", reference));

        cell.SetAttributeValue("t", isHeader ? "str" : DetermineValueType(value));
        if (value.StartsWith(" ", StringComparison.Ordinal) || value.EndsWith(" ", StringComparison.Ordinal))
        {
            cell.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        }

        var cellValue = new XElement(ns + XmlNames.Elements.CellValue, value);
        cell.Add(cellValue);
        return cell;
    }

    private static string DetermineValueType(string value)
    {
        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
        {
            return "n";
        }

        if (bool.TryParse(value, out _))
        {
            return "b";
        }

        return "str";
    }

    private static void IncrementAttribute(XElement element, string attributeName)
    {
        if (element.Attribute(attributeName) is XAttribute attr && int.TryParse(attr.Value, out var parsed))
        {
            attr.Value = (parsed + 1).ToString(CultureInfo.InvariantCulture);
        }
    }

    private static void SetElement(XDocument doc, XName parentName, XName elementName, string? value, XNamespace? xsi = null)
    {
        if (value is null)
        {
            return;
        }

        var parent = doc.Descendants(parentName).FirstOrDefault();
        if (parent is null)
        {
            return;
        }

        var element = parent.Element(elementName);
        if (element is null)
        {
            element = new XElement(elementName);
            parent.Add(element);
        }

        if (xsi is not null && elementName.NamespaceName.Contains("dcterms", StringComparison.Ordinal))
        {
            element.SetAttributeValue(xsi + "type", "dcterms:W3CDTF");
        }

        element.Value = value;
    }
}

