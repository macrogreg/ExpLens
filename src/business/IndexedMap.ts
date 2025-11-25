export class IndexedMap<TKey, TValue> {
    readonly #orderedData: TValue[] = [];
    readonly #indexedData = new Map<TKey, number>();

    get length(): number {
        return this.#orderedData.length;
    }

    tryAdd = (key: TKey, value: TValue) => {
        if (this.#indexedData.has(key)) {
            return false;
        }

        const p = this.#orderedData.push(value);
        this.#indexedData.set(key, p - 1);
        return true;
    };

    has = (key: TKey) => {
        return this.#indexedData.has(key);
    };

    getByIndex = (i: number) => {
        return this.#orderedData[i];
    };

    getByKey = (key: TKey) => {
        const p: number | undefined = this.#indexedData.get(key);
        if (p === undefined) {
            return undefined;
        }
        return this.getByIndex(p);
    };

    *[Symbol.iterator]() {
        yield* this.#orderedData;
    }
}
