// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import iconv from "iconv-lite";
import MashupHandler from "./mashupDocumentParser";
import WorkbookTemplate from "./workbookTemplate";

const pqCustomXmlPath: string = "customXml/item1.xml";
const connectionsXmlPath: string = "xl/connections.xml";
const queryTablesPath: string = "xl/queryTables/";
const pivotCachesPath: string = "xl/pivotCache/";

export class QueryInfo {
  queryMashup: string;
  refreshOnOpen: boolean;
  constructor(queryMashup: string, refreshOnOpen: boolean) {
    this.queryMashup = queryMashup;
    this.refreshOnOpen = refreshOnOpen;
  }
}
export class WorkbookManager {

  private mashupHandler: MashupHandler = new MashupHandler();

  async generateSingleQueryWorkbook(
    query: QueryInfo,
    templateFile?: File
  ): Promise<Blob> {
    let zip =
      templateFile === undefined
        ? await JSZip.loadAsync(WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE, {base64: true})
        : await JSZip.loadAsync(templateFile);

    return await this.generateSingleQueryWorkbookFromZip(zip, query);
  }

  private async generateSingleQueryWorkbookFromZip(
    zip: JSZip,
    query: QueryInfo
  ): Promise<Blob> {
    let old_base64 = await this.getBase64(zip);
    let new_base64 = await this.mashupHandler.ReplaceSingleQuery(
      old_base64,
      query.queryMashup
    );
    await this.setBase64(zip, new_base64);

    if (query.refreshOnOpen) {
      await this.setSingleQueryRefreshOnOpen(zip);
    }

    return await zip.generateAsync({
      type: "blob",
      mimeType: "application/xlsx",
    });
  }

  private async setSingleQueryRefreshOnOpen(zip: JSZip) {
    let connectionsXmlString = await zip.file(connectionsXmlPath)?.async("text");
    if (connectionsXmlString === undefined) {
      throw new Error("Connections were not found in template");
    }
    let parser: DOMParser = new DOMParser();
    let serializer = new XMLSerializer();

    let connectionsDoc: Document = parser.parseFromString(connectionsXmlString, "text/xml");
    let connectionId = "-1";
    let connectionsProperties = connectionsDoc.getElementsByTagName("dbPr");
    for (let properties of connectionsProperties){
      if (properties.getAttribute("command") == "SELECT * FROM [Query1]"){
        properties.parentElement?.setAttribute("refreshOnLoad","1");
        connectionId = properties.parentElement?.getAttribute("id")!;
        let newConn = serializer.serializeToString(connectionsDoc);
        zip.file(connectionsXmlPath, newConn);
        break;
      }
    }
    if (connectionId == "-1"){
      throw new Error("No connection found for Query1");
    }
    let found = false;

    // Find Query Table
    let queryTablePromises:Promise<{path:string,queryTableXmlString:string}>[] = [];

    zip.folder(queryTablesPath)?.forEach(async (relativePath,queryTableFile)=>{
      queryTablePromises.push((()=>{
        return queryTableFile.async("text").then(queryTableString=>{
          return {path:relativePath,queryTableXmlString:queryTableString};
        });
      })());
    });
    (await Promise.all(queryTablePromises)).forEach(({path, queryTableXmlString})=>{
      let queryTableDoc: Document = parser.parseFromString(queryTableXmlString, "text/xml");
      let element = queryTableDoc.getElementsByTagName("queryTable")[0];
      if(element.getAttribute("connectionId") == connectionId){
        element.setAttribute("refreshOnLoad","1");
        let newQT = serializer.serializeToString(queryTableDoc);
        zip.file(queryTablesPath+path, newQT);
        found=true;
      }
    })
    if (found){
      return;
    }

    // Find Query Table
    let pivotCachePromises:Promise<{path:string,pivotCacheXmlString:string}>[] = [];

    zip.folder(pivotCachesPath)?.forEach(async (relativePath,pivotCacheFile)=>{
      if (relativePath.startsWith("pivotCacheDefinition")){
        pivotCachePromises.push((()=>{
          return pivotCacheFile.async("text").then(pivotCacheString=>{
            return {path:relativePath,pivotCacheXmlString:pivotCacheString};
          });
        })());
      }
    });

    (await Promise.all(pivotCachePromises)).forEach(({path, pivotCacheXmlString})=>{
      let pivotCacheDoc: Document = parser.parseFromString(pivotCacheXmlString, "text/xml");
      let element = pivotCacheDoc.getElementsByTagName("cacheSource")[0];
      if(element.getAttribute("connectionId") == connectionId){
        element.parentElement!.setAttribute("refreshOnLoad","1");
        let newPC = serializer.serializeToString(pivotCacheDoc);
        zip.file(pivotCachesPath+path, newPC);
        found=true;
      }
    })
    if (!found){
      throw new Error("No Query Table or Pivot Table found for Query1 in given template.");
    }
  }

  private async setBase64(zip: JSZip, base64: string) {
    let newXml = `<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${base64}</DataMashup>`;
    let encoded = iconv.encode(newXml, "UCS2", { addBOM: true });
    zip.file(pqCustomXmlPath, encoded);
  }

  private async getBase64(zip: JSZip): Promise<string> {
    let xmlValue = await zip.file(pqCustomXmlPath)?.async("uint8array");
    if (xmlValue === undefined) {
      throw new Error("PQ document wasn't found in zip");
    }
    let xmlString = iconv.decode(xmlValue.buffer, "UTF-16");
    let parser: DOMParser = new DOMParser();
    let doc: Document = parser.parseFromString(xmlString, "text/xml");
    let result = doc.childNodes[0].textContent;
    if (result === null) {
      throw Error("Base64 wasn't found in zip");
    }
    return result;
  }
}