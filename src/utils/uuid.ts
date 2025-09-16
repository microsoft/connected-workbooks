// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { randomUUID } from "crypto";

export type UUIDGenerator = () => string;

let customUUIDGenerator: UUIDGenerator | null = null;

/**
 * Generates a UUID v4.
 *
 * Uses crypto.randomUUID() when available.
 * Falls back to a custom generator if one is set via setUUIDGenerator().
 *
 * @throws {Error} If no suitable UUID generator is available.
 */
export function generateUUID(): string {
    if (customUUIDGenerator) {
        return customUUIDGenerator();
    }

    if (typeof randomUUID === "function") {
        return randomUUID();
    }

    throw new Error("UUID generation not supported in this environment. Please provide a UUID generator via setUUIDGenerator()");
}

/**
 * Overrides the default UUID generator.
 *
 * This is useful in legacy environments where crypto.randomUUID() is not available.
 *
 * @param fn A function that returns a RFC4122-compliant UUID string.
 */
export function setUUIDGenerator(fn: UUIDGenerator): void {
    if (typeof fn !== "function") {
        throw new TypeError("UUID generator must be a function");
    }
    customUUIDGenerator = fn;
}