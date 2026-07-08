/**
 * In-memory mocks for Google Apps Script services used during Node.js tests.
 */

class ScriptCache {
    constructor() {
        /** @type {Map<string, string>} */
        this.store = new Map();
    }

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

    put(key, value, _seconds) {
        this.store.set(key, value);
    }

    putAll(obj, _seconds) {
        for (const [key, value] of Object.entries(obj)) {
            this.store.set(key, value);
        }
    }

    remove(key) {
        this.store.delete(key);
    }

    removeAll(keys) {
        for (const key of keys) {
            this.store.delete(key);
        }
    }

    clear() {
        this.store.clear();
    }
}

class ScriptProperties {
    constructor() {
        /** @type {Map<string, string>} */
        this.store = new Map();
    }

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

    setProperty(key, value) {
        this.store.set(key, value);
    }

    setProperties(obj) {
        for (const [key, value] of Object.entries(obj)) {
            this.store.set(key, value);
        }
    }

    deleteProperty(key) {
        this.store.delete(key);
    }

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
