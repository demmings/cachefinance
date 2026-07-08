/**
 * In-memory mocks for Google Apps Script services used during Node.js tests.
 */

/**
 * Copy every own enumerable entry from a plain object into a Map.
 * @param {Map<string, string>} store - Destination map.
 * @param {Record<string, string>} entries - Source key/value pairs.
 */
function copyEntriesInto(store, entries) {
    for (const [key, value] of Object.entries(entries)) {
        store.set(key, value);
    }
}

/**
 * In-memory mock for Apps Script script cache storage.
 */
class ScriptCache {
    constructor() {
        /** @type {Map<string, string>} */
        this.store = new Map();
    }

    /**
     * Read a single cache entry.
     * @param {string} key - Cache key to read.
     * @returns {string|null} Stored value, or null when missing.
     */
    get(key) {
        return this.store.has(key) ? this.store.get(key) : null;
    }

    getAll(keys) {
        /** @type {Record<string, string>} */
        const result = {};
        for (const key of keys) {
            if (this.store.has(key)) {
                result[key] = this.store.get(key);
            }
        }
        return result;
    }

    /**
     * Store a single cache entry.
     * @param {string} key - Cache key to set.
     * @param {string} value - Value to store.
     * @param {number} _seconds - Ignored; included for Apps Script API compatibility.
     */
    put(key, value, _seconds) {
        this.store.set(key, value);
    }

    /**
     * Store multiple cache entries at once.
     * @param {Record<string, string>} obj - Key/value pairs to store.
     * @param {number} _seconds - Ignored; included for Apps Script API compatibility.
     */
    putAll(obj, _seconds) {
        copyEntriesInto(this.store, obj);
    }

    /**
     * Remove a single cache entry.
     * @param {string} key - Cache key to remove.
     */
    remove(key) {
        this.store.delete(key);
    }

    /**
     * Remove multiple cache entries.
     * @param {string[]} keys - Cache keys to remove.
     */
    removeAll(keys) {
        for (const key of keys) {
            this.store.delete(key);
        }
    }

    /**
     * Remove all entries from the in-memory script cache.
     */
    clear() {
        this.store.clear();
    }
}

/**
 * In-memory mock for Apps Script script properties storage.
 */
class ScriptProperties {
    constructor() {
        /** @type {Map<string, string>} */
        this.store = new Map();
    }

    /**
     * Read a single script property.
     * @param {string} key - Property key to read.
     * @returns {string|null} Stored value, or null when missing.
     */
    getProperty(key) {
        return this.store.has(key) ? this.store.get(key) : null;
    }

    getProperties() {
        /** @type {Record<string, string>} */
        const result = {};
        for (const [key, value] of this.store.entries()) {
            result[key] = value;
        }
        return result;
    }

    /**
     * Store a single script property.
     * @param {string} key - Property key to set.
     * @param {string} value - Value to store.
     */
    setProperty(key, value) {
        this.store.set(key, value);
    }

    /**
     * Store multiple script properties at once.
     * @param {Record<string, string>} obj - Key/value pairs to store.
     */
    setProperties(obj) {
        copyEntriesInto(this.store, obj);
    }

    /**
     * Remove a single script property.
     * @param {string} key - Property key to delete.
     */
    deleteProperty(key) {
        this.store.delete(key);
    }

    /**
     * Remove all entries from the in-memory script properties store.
     */
    clear() {
        this.store.clear();
    }
}

const scriptCache = new ScriptCache();
const scriptProperties = new ScriptProperties();

export const CacheService = {
    getScriptCache() {
        return scriptCache;
    },
    reset() {
        scriptCache.clear();
    }
};

export const PropertiesService = {
    getScriptProperties() {
        return scriptProperties;
    },
    reset() {
        scriptProperties.clear();
    }
};

export const UrlFetchApp = {
  /** @type {Map<string, { content: string, throws?: Error }>} */
    responses: new Map(),

    fetch(url) {
        const response = this.responses.get(url);
        if (response?.throws) {
            throw response.throws;
        }

        return {
            getContentText() {
                return response?.content ?? "";
            }
        };
    },

    reset() {
        this.responses.clear();
    },

    mockResponse(url, content) {
        this.responses.set(url, { content });
    },

    fetchAll(requests) {
        return requests.map((request) => ({
            getContentText() {
                const response = UrlFetchApp.responses.get(request.url);
                return response?.content ?? "";
            }
        }));
    }
};

export const SpreadsheetApp = {
    getActiveSpreadsheet() {
        return {
            getRangeByName(_name) {
                return {
                    getValues() {
                        return [];
                    }
                };
            }
        };
    }
};
