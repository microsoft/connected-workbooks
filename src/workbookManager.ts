// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { pqUtils, documentUtils } from "./utils";
import WorkbookTemplate from "./workbookTemplate";
import MashupHandler from "./mashupDocumentParser";
import {
    connectionsXmlPath,
    queryTablesPath,
    pivotCachesPath,
    docPropsCoreXmlPath,
    defaultDocProps,
} from "./constants";
import {
    DocProps,
    QueryInfo,
    docPropsAutoUpdatedElements,
    docPropsModifiableElements,
} from "./types";


export class WorkbookManager {
    private mashupHandler: MashupHandler = new MashupHandler();

    async generateQuery1Workbook(
        query: QueryInfo,
        templateFile?: File,
        docProps?: DocProps
    ): Promise<Blob> {
        const zip =
            templateFile === undefined
                ? await JSZip.loadAsync(
                      WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE,
                      { base64: true }
                  )
                : await JSZip.loadAsync(templateFile);

        return await this.generateQuery1WorkbookFromZip(zip, query, docProps);
    }

    private async generateQuery1WorkbookFromZip(
        zip: JSZip,
        query: QueryInfo,
        docProps?: DocProps
    ): Promise<Blob> {
        await this.updatePowerQueryDocument(zip, query.queryMashup);
        await this.updateSingleQueryRefreshOnOpen(zip, query.refreshOnOpen);
        await this.updateDocProps(zip, docProps);

        return await zip.generateAsync({
            type: "blob",
            mimeType: "application/xlsx",
        });
    }

    private async updatePowerQueryDocument(zip: JSZip, queryMashup: string) {
        const mashupHandler = new MashupHandler();
        const old_base64 = await pqUtils.getBase64(zip);
        const new_base64 = await mashupHandler.ReplaceSingleQuery(
            old_base64,
            queryMashup
        );
        await pqUtils.setBase64(zip, new_base64);
    }

    private async updateDocProps(zip: JSZip, docProps: DocProps = {}) {

        //set defaults
        if (!docProps.title) docProps.title = defaultDocProps.title;
        if (!docProps.createdBy) docProps.createdBy = defaultDocProps.createdBy;
        if (!docProps.lastModifiedBy)
            docProps.lastModifiedBy = defaultDocProps.lastModifiedBy;

            //set auto updated elements
        const { doc, properties } = await documentUtils.getDocPropsProperties(
            zip
        );

        const docPropsAutoUpdatedElementsArr = Object.keys(
            docPropsAutoUpdatedElements
        ) as Array<keyof typeof docPropsAutoUpdatedElements>;

        docPropsAutoUpdatedElementsArr.forEach((tag) => {
            if (
                properties.getElementsByTagName(
                    docPropsAutoUpdatedElements[tag]
                ).length !== 1
            ) {
                throw new Error(
                    `Invalid DocProps core.xml - ${tag} does not appear exactly once.`
                );
            }
            documentUtils.createOrUpdateProperty(
                doc,
                properties,
                docPropsAutoUpdatedElements[tag],
                new Date().toISOString()
            );
        });

        //set modifiable elements
        const docPropsModifiableElementsArr = Object.keys(
            docPropsModifiableElements
        ) as Array<keyof typeof docPropsModifiableElements>;

        docPropsModifiableElementsArr
            .map((key) => ({
                    name: docPropsModifiableElements[key],
                    value: docProps[key],
            }))
            .forEach((kvp) => {
                documentUtils.createOrUpdateProperty(
                    doc,
                    properties,
                    kvp.name!,
                    kvp.value
                );
            });
            
        const serializer = new XMLSerializer();
        const newDoc = serializer.serializeToString(doc);
        zip.file(docPropsCoreXmlPath, newDoc);
    }

    private async updateSingleQueryRefreshOnOpen(
        zip: JSZip,
        refreshOnOpen: boolean
    ) {
        const connectionsXmlString = await zip
            .file(connectionsXmlPath)
            ?.async("text");
        if (connectionsXmlString === undefined) {
            throw new Error("Connections were not found in template");
        }
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();
        const refreshOnLoadValue = refreshOnOpen ? "1" : "0";
        const connectionsDoc: Document = parser.parseFromString(
            connectionsXmlString,
            "text/xml"
        );
        let connectionId = "-1";
        const connectionsProperties =
            connectionsDoc.getElementsByTagName("dbPr");

        const connectionsPropertiesArr = [...connectionsProperties];
        const queryProp = connectionsPropertiesArr.find((prop) => {
            prop.getAttribute("command") == "SELECT * FROM [Query1]";
        });
        if (queryProp) {
            queryProp.parentElement?.setAttribute(
                "refreshOnLoad",
                refreshOnLoadValue
            );
            const attr = queryProp.parentElement?.getAttribute("id");
            connectionId = attr!;
            const newConn = serializer.serializeToString(connectionsDoc);
            zip.file(connectionsXmlPath, newConn);
        }

        if (connectionId == "-1") {
            throw new Error("No connection found for Query1");
        }
        let found = false;

        // Find Query Table
        const queryTablePromises: Promise<{
            path: string;
            queryTableXmlString: string;
        }>[] = [];

        zip.folder(queryTablesPath)?.forEach(
            async (relativePath, queryTableFile) => {
                queryTablePromises.push(
                    (() => {
                        return queryTableFile
                            .async("text")
                            .then((queryTableString) => {
                                return {
                                    path: relativePath,
                                    queryTableXmlString: queryTableString,
                                };
                            });
                    })()
                );
            }
        );
        (await Promise.all(queryTablePromises)).forEach(
            ({ path, queryTableXmlString }) => {
                const queryTableDoc: Document = parser.parseFromString(
                    queryTableXmlString,
                    "text/xml"
                );
                const element =
                    queryTableDoc.getElementsByTagName("queryTable")[0];
                if (element.getAttribute("connectionId") == connectionId) {
                    element.setAttribute("refreshOnLoad", refreshOnLoadValue);
                    const newQT = serializer.serializeToString(queryTableDoc);
                    zip.file(queryTablesPath + path, newQT);
                    found = true;
                }
            }
        );
        if (found) {
            return;
        }

        // Find Pivot Table
        const pivotCachePromises: Promise<{
            path: string;
            pivotCacheXmlString: string;
        }>[] = [];

        zip.folder(pivotCachesPath)?.forEach(
            async (relativePath, pivotCacheFile) => {
                if (relativePath.startsWith("pivotCacheDefinition")) {
                    pivotCachePromises.push(
                        (() => {
                            return pivotCacheFile
                                .async("text")
                                .then((pivotCacheString) => {
                                    return {
                                        path: relativePath,
                                        pivotCacheXmlString: pivotCacheString,
                                    };
                                });
                        })()
                    );
                }
            }
        );
        (await Promise.all(pivotCachePromises)).forEach(
            ({ path, pivotCacheXmlString }) => {
                const pivotCacheDoc: Document = parser.parseFromString(
                    pivotCacheXmlString,
                    "text/xml"
                );
                const element =
                    pivotCacheDoc.getElementsByTagName("cacheSource")[0];
                if (element.getAttribute("connectionId") == connectionId) {
                    element.parentElement!.setAttribute(
                        "refreshOnLoad",
                        refreshOnLoadValue
                    );
                    const newPC = serializer.serializeToString(pivotCacheDoc);
                    zip.file(pivotCachesPath + path, newPC);
                    found = true;
                }
            }
        );
        if (!found) {
            throw new Error(
                "No Query Table or Pivot Table found for Query1 in given template."
            );
        }
    }
}
