"use strict";
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    Object.defineProperty(o, k2, { enumerable: true, get: function() { return m[k]; } });
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
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
const jszip_1 = __importDefault(require("jszip"));
const base64 = __importStar(require("byte-base64"));
const arrayUtils_1 = require("./arrayUtils");
class MashupHandler {
    ReplaceSingleQuery(base64str, query) {
        return __awaiter(this, void 0, void 0, function* () {
            let buffer = base64.base64ToBytes(base64str).buffer;
            let mashupArray = new arrayUtils_1.ArrayReader(buffer);
            let startArray = mashupArray.getBytes(4);
            let packageSize = mashupArray.getInt32();
            let packageOPC = mashupArray.getBytes(packageSize);
            let endBuffer = mashupArray.getBytes();
            let newPackageBuffer = yield this.editSingleQueryPackage(packageOPC, "Query1", query);
            let packageSizeBuffer = arrayUtils_1.getInt32Buffer(newPackageBuffer.byteLength);
            let newMashup = arrayUtils_1.concatArrays(startArray, packageSizeBuffer, newPackageBuffer, endBuffer);
            return base64.bytesToBase64(newMashup);
        });
    }
    editSingleQueryPackage(packageOPC, queryName, query) {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            let packageZip = yield jszip_1.default.loadAsync(packageOPC);
            let section1m = yield ((_a = packageZip.file("Formulas/Section1.m")) === null || _a === void 0 ? void 0 : _a.async("text"));
            if (section1m === undefined) {
                throw new Error("Formula section wasn't found in template");
            }
            let newSection1m = `section Section1;
    
    shared ${queryName} = 
    ${query};`;
            packageZip.file("Formulas/Section1.m", newSection1m, { compression: "" });
            return yield packageZip.generateAsync({ type: "uint8array" });
        });
    }
}
exports.default = MashupHandler;
//# sourceMappingURL=mashupDocumentParser.js.map