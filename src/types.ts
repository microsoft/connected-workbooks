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
    columnNames: string[];
    rows: string[][];
}

export interface Grid {
    data: (string | number | boolean)[][];
    config?: GridConfig;
}

export interface GridConfig {
    promoteHeaders?: boolean;
    adjustColumnNames?: boolean;
}

export interface FileConfigs {
    templateFile?: File;
    docProps?: DocProps;
}

export interface DateTimeFormat {
    regex: RegExp;
    validateDate: (data: string) => boolean;
    isDate: boolean;
    formatCode?: string;
    formatId?: number;
}

export enum DataTypes {
    null = 0,
    string = 1,
    number = 2,
    boolean = 3,
    dateTime = 4,
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
