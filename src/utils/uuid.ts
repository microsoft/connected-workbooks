// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export type UUIDGenerator = () => string;

let customUUIDGenerator: UUIDGenerator | null = null;
let nodeRandomUUID: (() => string) | null = null;

// Try to get Node.js crypto.randomUUID at module load time
if (typeof process !== "undefined" && process.versions?.node) {
    try {
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const crypto = require("crypto");
        if (typeof crypto.randomUUID === "function") {
            nodeRandomUUID = crypto.randomUUID.bind(crypto);
        }
    } catch {
        // crypto module not available
    }
}
/**
 * Generates a UUID v4.
 *
 * Requirements:
 * - Browser: Chrome 92+, Firefox 95+, Safari 15.4+ (requires secure context)
 * - Node.js: 14.17.0+
 * - Or provide a custom generator via setUUIDGenerator()
 */
export function generateUUID(): string {
    // Custom: use custom generator if set
    if (customUUIDGenerator) {
        return customUUIDGenerator();
    }

    // Modern environments have crypto.randomUUID
    if (typeof globalThis.crypto?.randomUUID === "function") {
        return globalThis.crypto.randomUUID();
    }

    if (nodeRandomUUID) {
        return nodeRandomUUID();
    }

    throw new Error(
        "UUID generation not supported in this environment. " +
        "Requires Node.js >= 14.17.0 or modern browser with secure context. " +
        "Alternatively, provide a UUID generator via setUUIDGenerator()."
    );
}

/**
 * Overrides the default UUID generator.
 * Useful in legacy environments where no native support exists.
 */
export function setUUIDGenerator(fn: UUIDGenerator): void {
    if (typeof fn !== "function") {
        throw new TypeError("UUID generator must be a function");
    }
    customUUIDGenerator = fn;
}

/**
 * Clears a previously set custom UUID generator, returning to environment detection.
 */
export function clearUUIDGenerator(): void {
    customUUIDGenerator = null;
}