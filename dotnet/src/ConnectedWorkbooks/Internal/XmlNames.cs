// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class XmlNames
{
    internal static class Elements
    {
        public const string SharedStringTable = "sst";
        public const string SharedStringItem = "si";
        public const string Text = "t";
        public const string CellValue = "v";
        public const string DatabaseProperties = "dbPr";
        public const string QueryTable = "queryTable";
        public const string CacheSource = "cacheSource";
        public const string Table = "table";
        public const string TableColumns = "tableColumns";
        public const string TableColumn = "tableColumn";
        public const string AutoFilter = "autoFilter";
        public const string SheetData = "sheetData";
        public const string Row = "row";
        public const string Cell = "c";
        public const string DefinedName = "definedName";
        public const string QueryTableFields = "queryTableFields";
        public const string QueryTableField = "queryTableField";
        public const string QueryTableRefresh = "queryTableRefresh";
        public const string Relationships = "Relationships";
        public const string Relationship = "Relationship";
        public const string Item = "Item";
        public const string ItemPath = "ItemPath";
        public const string Entry = "Entry";
        public const string Items = "Items";
        public const string Worksheet = "worksheet";
        public const string Dimension = "dimension";
        public const string Selection = "selection";
    }

    internal static class Attributes
    {
        public const string Count = "count";
        public const string UniqueCount = "uniqueCount";
        public const string RefreshOnLoad = "refreshOnLoad";
        public const string ConnectionId = "connectionId";
        public const string Name = "name";
        public const string Description = "description";
        public const string Connection = "connection";
        public const string Command = "command";
        public const string Id = "id";
        public const string RelId = "r:id";
        public const string Target = "Target";
        public const string PartName = "PartName";
        public const string ContentType = "ContentType";
        public const string Reference = "ref";
        public const string SqRef = "sqref";
        public const string TableColumnId = "tableColumnId";
        public const string UniqueName = "uniqueName";
        public const string QueryTableFieldId = "queryTableFieldId";
        public const string NextId = "nextId";
        public const string Row = "r";
        public const string Spans = "spans";
        public const string X14acDyDescent = "x14ac:dyDescent";
        public const string Type = "Type";
        public const string Value = "Value";
        public const string ResultType = "ResultType";
    }
}

