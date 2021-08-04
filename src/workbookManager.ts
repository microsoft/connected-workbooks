// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import MashupHandler from "./mashupDocumentParser";
import { zipUtils, JSZip } from "./utils";
import {
    connectionsXmlPath,
    queryTablesPath,
    pivotCachesPath,
} from "./constants";

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
        const zip =
            templateFile === undefined
                ? await zipUtils.loadAsyncDefaultTemplate()
                : await zipUtils.loadAsync(templateFile);

        return await this.generateSingleQueryWorkbookFromZip(zip, query);
    }

    private async generateSingleQueryWorkbookFromZip(
        zip: JSZip,
        query: QueryInfo
    ): Promise<Blob> {
        const old_base64 = await zipUtils.getBase64(zip);
        const new_base64 = await this.mashupHandler.ReplaceSingleQuery(
            old_base64,
            query.queryMashup
        );
        await zipUtils.setBase64(zip, new_base64);

        if (query.refreshOnOpen) {
            await this.setSingleQueryRefreshOnOpen(zip);
        }

        return await zip.generateAsync({
            type: "blob",
            mimeType: "application/xlsx",
        });
    }

    private async setSingleQueryRefreshOnOpen(zip: JSZip) {
        const connectionsXmlString = await zip
            .file(connectionsXmlPath)
            ?.async("text");
        if (connectionsXmlString === undefined) {
            throw new Error("Connections were not found in template");
        }
        const parser: DOMParser = new DOMParser();
        const serializer = new XMLSerializer();

        const connectionsDoc: Document = parser.parseFromString(
            connectionsXmlString,
            "text/xml"
        );
        let connectionId = "-1";
        const connectionsProperties =
            connectionsDoc.getElementsByTagName("dbPr");
        for (const properties of connectionsProperties) {
            if (
                properties.getAttribute("command") == "SELECT * FROM [Query1]"
            ) {
                properties.parentElement?.setAttribute("refreshOnLoad", "1");
                const attr = properties.parentElement?.getAttribute("id");
                connectionId = attr!;
                const newConn = serializer.serializeToString(connectionsDoc);
                zip.file(connectionsXmlPath, newConn);
                break;
            }
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
                    element.setAttribute("refreshOnLoad", "1");
                    const newQT = serializer.serializeToString(queryTableDoc);
                    zip.file(queryTablesPath + path, newQT);
                    found = true;
                }
            }
        );
        if (found) {
            return;
        }

        // Find Query Table
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
                    element.parentElement!.setAttribute("refreshOnLoad", "1");
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
