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