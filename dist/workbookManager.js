"use strict";
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.WorkbookManager = exports.QueryInfo = void 0;
const jszip_1 = __importDefault(require("jszip"));
const iconv_lite_1 = __importDefault(require("iconv-lite"));
const mashupDocumentParser_1 = __importDefault(require("./mashupDocumentParser"));
const workbookTemplate_1 = __importDefault(require("./workbookTemplate"));
const pqCustomXmlPath = "customXml/item1.xml";
const connectionsXmlPath = "xl/connections.xml";
const queryTablesPath = "xl/queryTables/";
const pivotCachesPath = "xl/pivotCache/";
class QueryInfo {
    constructor(queryMashup, refreshOnOpen) {
        this.queryMashup = queryMashup;
        this.refreshOnOpen = refreshOnOpen;
    }
}
exports.QueryInfo = QueryInfo;
class WorkbookManager {
    constructor() {
        this.mashupHandler = new mashupDocumentParser_1.default();
    }
    generateSingleQueryWorkbook(query, templateFile) {
        return __awaiter(this, void 0, void 0, function* () {
            let zip = templateFile === undefined
                ? yield jszip_1.default.loadAsync(workbookTemplate_1.default.SIMPLE_QUERY_WORKBOOK_TEMPLATE, { base64: true })
                : yield jszip_1.default.loadAsync(templateFile);
            return yield this.generateSingleQueryWorkbookFromZip(zip, query);
        });
    }
    generateSingleQueryWorkbookFromZip(zip, query) {
        return __awaiter(this, void 0, void 0, function* () {
            let old_base64 = yield this.getBase64(zip);
            let new_base64 = yield this.mashupHandler.ReplaceSingleQuery(old_base64, query.queryMashup);
            yield this.setBase64(zip, new_base64);
            if (query.refreshOnOpen) {
                yield this.setSingleQueryRefreshOnOpen(zip);
            }
            return yield zip.generateAsync({
                type: "blob",
                mimeType: "application/xlsx",
            });
        });
    }
    setSingleQueryRefreshOnOpen(zip) {
        var _a, _b, _c, _d, _e;
        return __awaiter(this, void 0, void 0, function* () {
            let connectionsXmlString = yield ((_a = zip.file(connectionsXmlPath)) === null || _a === void 0 ? void 0 : _a.async("text"));
            if (connectionsXmlString === undefined) {
                throw new Error("Connections were not found in template");
            }
            let parser = new DOMParser();
            let serializer = new XMLSerializer();
            let connectionsDoc = parser.parseFromString(connectionsXmlString, "text/xml");
            let connectionId = "-1";
            let connectionsProperties = connectionsDoc.getElementsByTagName("dbPr");
            for (let properties of connectionsProperties) {
                if (properties.getAttribute("command") == "SELECT * FROM [Query1]") {
                    (_b = properties.parentElement) === null || _b === void 0 ? void 0 : _b.setAttribute("refreshOnLoad", "1");
                    connectionId = (_c = properties.parentElement) === null || _c === void 0 ? void 0 : _c.getAttribute("id");
                    let newConn = serializer.serializeToString(connectionsDoc);
                    console.log("newConn:", newConn);
                    zip.file(connectionsXmlPath, newConn);
                    break;
                }
            }
            if (connectionId == "-1") {
                throw new Error("No connection found for Query1");
            }
            let found = false;
            // Find Query Table
            let queryTablePromises = [];
            (_d = zip.folder(queryTablesPath)) === null || _d === void 0 ? void 0 : _d.forEach((relativePath, queryTableFile) => __awaiter(this, void 0, void 0, function* () {
                queryTablePromises.push((() => {
                    return queryTableFile.async("text").then(queryTableString => {
                        return { path: relativePath, queryTableXmlString: queryTableString };
                    });
                })());
            }));
            (yield Promise.all(queryTablePromises)).forEach(({ path, queryTableXmlString }) => {
                let queryTableDoc = parser.parseFromString(queryTableXmlString, "text/xml");
                let element = queryTableDoc.getElementsByTagName("queryTable")[0];
                console.log(element.getAttribute("connectionId"));
                console.log(connectionId);
                if (element.getAttribute("connectionId") == connectionId) {
                    element.setAttribute("refreshOnLoad", "1");
                    let newQT = serializer.serializeToString(queryTableDoc);
                    zip.file(queryTablesPath + path, newQT);
                    found = true;
                }
            });
            if (found) {
                return;
            }
            // Find Query Table
            let pivotCachePromises = [];
            console.log("looking for cache");
            (_e = zip.folder(pivotCachesPath)) === null || _e === void 0 ? void 0 : _e.forEach((relativePath, pivotCacheFile) => __awaiter(this, void 0, void 0, function* () {
                console.log(relativePath);
                if (relativePath.startsWith("pivotCacheDefinition")) {
                    console.log("Found pivot cache");
                    pivotCachePromises.push((() => {
                        return pivotCacheFile.async("text").then(pivotCacheString => {
                            return { path: relativePath, pivotCacheXmlString: pivotCacheString };
                        });
                    })());
                }
            }));
            (yield Promise.all(pivotCachePromises)).forEach(({ path, pivotCacheXmlString }) => {
                let pivotCacheDoc = parser.parseFromString(pivotCacheXmlString, "text/xml");
                let element = pivotCacheDoc.getElementsByTagName("cacheSource")[0];
                console.log(element.getAttribute("connectionId"));
                console.log(connectionId);
                if (element.getAttribute("connectionId") == connectionId) {
                    element.parentElement.setAttribute("refreshOnLoad", "1");
                    let newPC = serializer.serializeToString(pivotCacheDoc);
                    zip.file(pivotCachesPath + path, newPC);
                    found = true;
                }
            });
            if (!found) {
                throw new Error("No Query Table or Pivot Table found for Query1 in given template.");
            }
        });
    }
    setBase64(zip, base64) {
        return __awaiter(this, void 0, void 0, function* () {
            let newXml = `<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${base64}</DataMashup>`;
            let encoded = iconv_lite_1.default.encode(newXml, "UCS2", { addBOM: true });
            zip.file(pqCustomXmlPath, encoded);
        });
    }
    getBase64(zip) {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            let xmlValue = yield ((_a = zip.file(pqCustomXmlPath)) === null || _a === void 0 ? void 0 : _a.async("uint8array"));
            if (xmlValue === undefined) {
                throw new Error("PQ document wasn't found in zip");
            }
            let xmlString = iconv_lite_1.default.decode(xmlValue.buffer, "UTF-16");
            let parser = new DOMParser();
            let doc = parser.parseFromString(xmlString, "text/xml");
            let result = doc.childNodes[0].textContent;
            if (result === null) {
                throw Error("Base64 wasn't found in zip");
            }
            return result;
        });
    }
}
exports.WorkbookManager = WorkbookManager;
//# sourceMappingURL=workbookManager.js.map