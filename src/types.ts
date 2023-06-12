// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface QueryInfo {
    refreshOnOpen: boolean;
    queryMashup: string;
    queryName?: string;
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
    columnMetadata: ColumnMetadata[];
    rows: string[][];
    columnwidth?: number;
}

export interface ColumnMetadata {
    name: string;
    type: DataTypes;
}

export interface Grid {
    header: ColumnMetadata[];
    gridData: (string|number|boolean)[][];
}

export interface TableDataParser {
    parseToTableData: (grid: any) => TableData | undefined;
}

export enum DataTypes {
    autodetect = -1,
    null = 0,
    string = 1,
    number = 2,
    boolean = 3,
}

export enum DocPropsModifiableElements {
    title = "dc:title",
    subject = "dc:subject",
    keywords = "cp:keywords",
    createdBy = "dc:creator",
    description = "dc:description",
    lastModifiedBy = "cp:lastModifiedBy",
    category = "cp:category",
    revision = "cp:revision",
}

export enum DocPropsAutoUpdatedElements {
    created = "dcterms:created",
    modified = "dcterms:modified",
}
