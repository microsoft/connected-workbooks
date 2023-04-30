// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface QueryInfo {
    queryMashup: string;
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

export interface TableData {
    columnMetadata: ColumnMetadata[];
    data: string[][];
    columnwidth?: number;
}

export interface ColumnMetadata {
    name: string;
    type: number;
}

export interface Grid {
    Header: ColumnMetadata[];
    GridData: dataTypes[][];
}

export interface TableDataParser {
    parseToTableData: (grid: any) => TableData | undefined;
}

export enum dataTypes {
    null = 0,
    string = 1,
    number = 2,
    boolean = 3,
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
