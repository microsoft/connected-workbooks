// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { types } from "@babel/core";

export interface QueryInfo {
    queryMashup: string;
    queryName?: string;
    refreshOnOpen: boolean;
}

export interface DocProps {
    title?: string | null;
    subject?: string | null;
    keywords?: string | null;
    createdBy?: string | null;
    description?: string | null;
    lastModifiedBy?: string | null;
    category?: string | null;
    revision?: string | null;
}

export interface Metadata {
    queryName: string;
}

export interface TableData {
    columnNames: string[];
    columnTypes: number[];
    columnwidth?: number;
    data: string[][];
}

export enum dataTypes {
    null = 0,
    string = 1,
    number = 2,
    boolean = 3
}

export enum docPropsModifiableElements {
    title = "dc:title",
    subject = "dc:subject",
    keywords = "cp:keywords",
    createdBy = "dc:creator",
    description = "dc:description",
    lastModifiedBy = "cp:lastModifiedBy",
    category = "cp:category",
    revision = "cp:revision",
}

export enum docPropsAutoUpdatedElements {
    created = "dcterms:created",
    modified = "dcterms:modified",
}
