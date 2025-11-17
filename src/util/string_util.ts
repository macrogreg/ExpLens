export function containsAtLeastNRegex(haystack: string, needle: RegExp, n: number): boolean {
    if (n < 0) {
        return false;
    }

    if (n === 0) {
        return true;
    }

    if (!needle.global) {
        needle = new RegExp(needle.source, needle.flags + "g");
    }

    let count = 0;
    let match;
    while ((match = needle.exec(haystack)) !== null) {
        count++;
        if (count >= n) {
            return true;
        }

        // avoid infinite loop on zero-width matches:
        if (match.index === needle.lastIndex) {
            needle.lastIndex++;
        }
    }

    return false;
}

export function containsAtLeastN(haystack: string, needle: string, n: number): boolean {
    if (n < 0) {
        return false;
    }

    if (n === 0) {
        return true;
    }

    if (!haystack || !needle || haystack.length === 0 || needle.length === 0) {
        return false;
    }

    let count = 0;
    let pos = 0;
    while ((pos = haystack.indexOf(needle, pos)) !== -1) {
        count++;
        if (count >= n) {
            return true;
        }
        pos += needle.length;
    }

    return false;
}

// eslint-disable-next-line @typescript-eslint/no-wrapper-object-types
export function isNullOrWhitespace(s: unknown): s is null | undefined | string | String {
    if (s === null) return true;
    if (s === undefined) return true;
    if (typeof s === "string" || s instanceof String) {
        if (s.length === 0) return true;
        return s.trim().length === 0;
    }
    return false;
}

// eslint-disable-next-line @typescript-eslint/no-wrapper-object-types
export function isNotNullOrWhitespaceStr(s: unknown): s is string | String {
    if (s === null) return false;
    if (s === undefined) return false;
    if (typeof s === "string" || s instanceof String) {
        if (s.length === 0) return false;
        return s.trim().length > 0;
    }
    return false;
}

export function strcmp(a: string | null | undefined, b: string | null | undefined): number {
    if (a === b) return 0;

    const isANil = a === null || a === undefined;
    const isBNil = b === null || b === undefined;

    if (isANil && isBNil) return 0;
    if (isANil) return -1;
    if (isBNil) return 1;

    return a < b ? -1 : a > b ? 1 : 0;
}
