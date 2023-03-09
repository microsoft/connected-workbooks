export const connectedWorkbookXmlMock =
    '<?xml version="1.0" encoding="utf-8"?><ConnectedWorkbook xmlns="http://schemas.microsoft.com/ConnectedWorkbook" version="1.0.0"></ConnectedWorkbook>';

export const sharedStringsXmlMock =
`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>Query1</t></si><si><t/></si><si><t>newQueryName</t></si></sst>`

export const existingSharedStringsXmlMock =
`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>newQueryName</t></si><si><t/></si></sst>`

export const pqMetadataXmlMockPart1 = '<LocalPackageMetadataFile xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><Items>    <Item>      <ItemLocation>        <ItemType>AllFormulas</ItemType>        <ItemPath/>      </ItemLocation>      <StableEntries/>    </Item>    <Item>      <ItemLocation>        <ItemType>Formula</ItemType>        <ItemPath>Section1/newQueryName</ItemPath>      </ItemLocation>      <StableEntries>        <Entry Type="IsPrivate" Value="l0"/>        <Entry Type="FillEnabled" Value="l1"/>        <Entry Type="FillObjectType" Value="sTable"/>        <Entry Type="FillToDataModelEnabled" Value="l0"/>       <Entry Type="BufferNextRefresh" Value="l1"/>        <Entry Type="ResultType" Value="sTable"/>        <Entry Type="NameUpdatedAfterFill" Value="l0"/>        <Entry Type="NavigationStepName" Value="sNavigation"/>        <Entry Type="FillTarget" Value="snewQueryName"/>        <Entry Type="FilledCompleteResultToWorksheet" Value="l1"/>        <Entry Type="AddedToDataModel" Value="l0"/>        <Entry Type="FillCount" Value="l1"/>        <Entry Type="FillErrorCode" Value="sUnknown"/>        <Entry Type="FillErrorCount" Value="l0"/>'
export const pqMetadataXmlMockPart2 = '<Entry Type="FillColumnTypes" Value="sBg=="/>        <Entry Type="FillColumnNames" Value="s[&quot;newQueryName&quot;]"/>        <Entry Type="FillStatus" Value="sComplete"/>        <Entry Type="RelationshipInfoContainer" Value="s{&quot;columnCount&quot;:1,&quot;keyColumnNames&quot;:[],&quot;queryRelationships&quot;:[],&quot;columnIdentities&quot;:[&quot;Section1/newQueryName/AutoRemovedColumns1.{newQueryName,0}&quot;],&quot;ColumnCount&quot;:1,&quot;KeyColumnNames&quot;:[],&quot;ColumnIdentities&quot;:[&quot;Section1/newQueryName/AutoRemovedColumns1.{newQueryName,0}&quot;],&quot;RelationshipInfo&quot;:[]}"/>      </StableEntries>    </Item>    <Item>      <ItemLocation>        <ItemType>Formula</ItemType>        <ItemPath>Section1/newQueryName/Source</ItemPath>      </ItemLocation>      <StableEntries/>    </Item>  </Items> </LocalPackageMetadataFile>'

export const pqMetadataXmlMock = `
<LocalPackageMetadataFile xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Items>
    <Item>
      <ItemLocation>
        <ItemType>AllFormulas</ItemType>
        <ItemPath />
      </ItemLocation>
      <StableEntries />
    </Item>
    <Item>
      <ItemLocation>
        <ItemType>Formula</ItemType>
        <ItemPath>Section1/Query1</ItemPath>
      </ItemLocation>
      <StableEntries>
        <Entry Type="IsPrivate" Value="l0" />
        <Entry Type="FillEnabled" Value="l1" />
        <Entry Type="FillObjectType" Value="sTable" />
        <Entry Type="FillToDataModelEnabled" Value="l0" />
        <Entry Type="BufferNextRefresh" Value="l1" />
        <Entry Type="ResultType" Value="sTable" />
        <Entry Type="NameUpdatedAfterFill" Value="l0" />
        <Entry Type="FillTarget" Value="sQuery1" />
        <Entry Type="FilledCompleteResultToWorksheet" Value="l1" />
        <Entry Type="AddedToDataModel" Value="l0" />
        <Entry Type="FillCount" Value="l1" />
        <Entry Type="FillErrorCode" Value="sUnknown" />
        <Entry Type="FillErrorCount" Value="l0" />
        <Entry Type="FillLastUpdated" Value="d2023-03-06T08:04:33.3178365Z" />
        <Entry Type="FillColumnTypes" Value="sBg==" />
        <Entry Type="FillColumnNames" Value="s[&quot;Column1&quot;]" />
        <Entry Type="FillStatus" Value="sComplete" />
        <Entry Type="RelationshipInfoContainer" Value="s{&quot;columnCount&quot;:1,&quot;keyColumnNames&quot;:[],&quot;queryRelationships&quot;:[],&quot;columnIdentities&quot;:[&quot;Section1/Query1/AutoRemovedColumns1.{Column1,0}&quot;],&quot;ColumnCount&quot;:1,&quot;KeyColumnNames&quot;:[],&quot;ColumnIdentities&quot;:[&quot;Section1/Query1/AutoRemovedColumns1.{Column1,0}&quot;],&quot;RelationshipInfo&quot;:[]}" />
      </StableEntries>
    </Item>
    <Item>
      <ItemLocation>
        <ItemType>Formula</ItemType>
        <ItemPath>Section1/Query1/Source</ItemPath>
      </ItemLocation>
      <StableEntries />
    </Item>
  </Items>
</LocalPackageMetadataFile>`