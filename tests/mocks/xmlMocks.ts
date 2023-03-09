export const connectedWorkbookXmlMock =
    '<?xml version="1.0" encoding="utf-8"?><ConnectedWorkbook xmlns="http://schemas.microsoft.com/ConnectedWorkbook" version="1.0.0"></ConnectedWorkbook>';

export const sharedStringsXmlMock =
`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>Query1</t></si><si><t/></si><si><t>newQueryName</t></si></sst>`

export const existingSharedStringsXmlMock =
`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>newQueryName</t></si><si><t/></si></sst>`

export const queryTablesXmlMock = 
    `<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16" mc:Ignorable="xr16" name="ExternalData_1" connectionId="2" xr16:uid="{288FEC94-24B8-497A-B2F3-2E025A1F19F1}" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0">
<queryTableRefresh nextId="2">
<queryTableFields count="1">
<queryTableField id="1" name="Column1" tableColumnId="1"/>
</queryTableFields>
</queryTableRefresh>
</queryTable>`
export const pqMetadataXmlMockPart1 = '<LocalPackageMetadataFile xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><Items>    <Item>      <ItemLocation>        <ItemType>AllFormulas</ItemType>        <ItemPath/>      </ItemLocation>      <StableEntries/>    </Item>    <Item>      <ItemLocation>        <ItemType>Formula</ItemType>        <ItemPath>Section1/newQueryName</ItemPath>      </ItemLocation>      <StableEntries>        <Entry Type="IsPrivate" Value="l0"/>        <Entry Type="FillEnabled" Value="l1"/>        <Entry Type="FillObjectType" Value="sTable"/>        <Entry Type="FillToDataModelEnabled" Value="l0"/>       <Entry Type="BufferNextRefresh" Value="l1"/>        <Entry Type="ResultType" Value="sTable"/>        <Entry Type="NameUpdatedAfterFill" Value="l0"/>        <Entry Type="NavigationStepName" Value="sNavigation"/>        <Entry Type="FillTarget" Value="snewQueryName"/>        <Entry Type="FilledCompleteResultToWorksheet" Value="l1"/>        <Entry Type="AddedToDataModel" Value="l0"/>        <Entry Type="FillCount" Value="l1"/>        <Entry Type="FillErrorCode" Value="sUnknown"/>        <Entry Type="FillErrorCount" Value="l0"/>'
export const pqMetadataXmlMockPart2 = '<Entry Type="FillColumnTypes" Value="sBg=="/>        <Entry Type="FillColumnNames" Value="s[&quot;newQueryName&quot;]"/>        <Entry Type="FillStatus" Value="sComplete"/>        <Entry Type="RelationshipInfoContainer" Value="s{&quot;columnCount&quot;:1,&quot;keyColumnNames&quot;:[],&quot;queryRelationships&quot;:[],&quot;columnIdentities&quot;:[&quot;Section1/newQueryName/AutoRemovedColumns1.{newQueryName,0}&quot;],&quot;ColumnCount&quot;:1,&quot;KeyColumnNames&quot;:[],&quot;ColumnIdentities&quot;:[&quot;Section1/newQueryName/AutoRemovedColumns1.{newQueryName,0}&quot;],&quot;RelationshipInfo&quot;:[]}"/>      </StableEntries>    </Item>    <Item>      <ItemLocation>        <ItemType>Formula</ItemType>        <ItemPath>Section1/newQueryName/Source</ItemPath>      </ItemLocation>      <StableEntries/>    </Item>  </Items> </LocalPackageMetadataFile>'
