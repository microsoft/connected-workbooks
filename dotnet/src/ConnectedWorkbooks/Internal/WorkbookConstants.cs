// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Text.RegularExpressions;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class WorkbookConstants
{
    public const string ConnectionsXmlPath = "xl/connections.xml";
    public const string SharedStringsXmlPath = "xl/sharedStrings.xml";
    public const string DefaultSheetPath = "xl/worksheets/sheet1.xml";
    public const string DefaultTablePath = "xl/tables/table1.xml";
    public const string QueryTablesFolder = "xl/queryTables/";
    public const string QueryTablePath = "xl/queryTables/queryTable1.xml";
    public const string WorkbookXmlPath = "xl/workbook.xml";
    public const string WorkbookRelsPath = "xl/_rels/workbook.xml.rels";
    public const string PivotCachesFolder = "xl/pivotCache/";
    public const string Section1mPath = "Formulas/Section1.m";
    public const string DocPropsCoreXmlPath = "docProps/core.xml";
    public const string DocPropsAppXmlPath = "docProps/app.xml";
    public const string ContentTypesPath = "[Content_Types].xml";
    public const string RootRelsPath = "_rels/.rels";
    public const string DocMetadataPath = "docMetadata";
    public const string CustomXmlFolder = "customXml";
    public const string LabelInfoPath = "docMetadata/LabelInfo.xml";
    public const string TablesFolder = "xl/tables/";

    public const string ConnectedWorkbookNamespace = "http://schemas.microsoft.com/ConnectedWorkbook";
    public const string DataMashupNamespace = "http://schemas.microsoft.com/DataMashup";

    public const int MaxQueryLength = 80;
    public const int MaxCellCharacters = 32767;

    public static readonly Regex CustomXmlItemRegex = new(@"^customXml/item(\d+)\.xml$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
}

