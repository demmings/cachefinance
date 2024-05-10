/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { ScriptSettings, PropertyData } from "./SQL/ScriptSettings.js";
export { CacheFinanceUtils };

class Logger {
    static log(msg) {
        console.log(msg);
    }
}
//  *** DEBUG END ***/

/**
 * @classdesc Multi-purpose functions used within the cache finance custom functions.
 */
class CacheFinanceUtils {
    /**
     * 
     * @param {String[]} symbols 
     * @param {String} attribute 
     */
    static bulkShortCacheRemoveAll(symbols, attribute) {
        const cacheKeyList = CacheFinanceUtils.createCacheKeyList(symbols, attribute);
        CacheService.getScriptCache().removeAll(cacheKeyList);
    }

    /**
     * 
     * @param {any[]} symbols 
     * @param {String} attribute 
     * @returns {any[]} 
     */
    static bulkShortCacheGet(symbols, attribute) {
        const cacheKeyList = CacheFinanceUtils.createCacheKeyList(symbols, attribute);
        return CacheFinanceUtils.getFinanceValuesFromShortCache(cacheKeyList);
    }

    /**
     * 
     * @param {any[]} symbols 
     * @param {String} attribute 
     * @param {any[]} newCacheData 
     */
    static bulkShortCachePut(symbols, attribute, newCacheData, cacheSeconds) {
        if (symbols.length === 0) {
            return;
        }

        const cacheKeyList = CacheFinanceUtils.createCacheKeyList(symbols, attribute);
        CacheFinanceUtils.putFinanceValuesIntoShortCache(cacheKeyList, newCacheData, cacheSeconds);
    }

    /**
     * 
     * @param {any[]} symbols 
     * @param {String} attribute 
     * @param {any[]} cacheData 
     */
    static bulkLongCachePut(symbols, attribute, cacheData, daysToHold=7) {
        const cacheKeys = CacheFinanceUtils.createCacheKeyList(symbols, attribute);
        const newCacheKeys = [];
        const newCacheData = [];

        for (let i = 0; i < cacheKeys.length; i++) {
            if (CacheFinanceUtils.isValidGoogleValue(cacheData[i])) {
                newCacheKeys.push(cacheKeys[i]);
                newCacheData.push(cacheData[i])
            }
        }

        ScriptSettings.putAllKeysWithData(newCacheKeys, newCacheData, daysToHold);
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {String} attribute 
     * @returns {any[]}
     */
    static bulkLongCacheGet(symbols, attribute) {
        const cacheKeyList = CacheFinanceUtils.createCacheKeyList(symbols, attribute);
        return ScriptSettings.getAll(cacheKeyList);
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {String} attribute 
     * @returns {String[]}
     */
    static createCacheKeyList(symbols, attribute) {
        return symbols.map(symbol => CacheFinanceUtils.makeCacheKey(symbol.toUpperCase(), attribute));
    }

    /**
     * 
     * @param {String[]} cacheKeys 
     * @returns {any[]} - A missing cache entry will return 'null'
     */
    static getFinanceValuesFromShortCache(cacheKeys) {
        const shortCache = CacheService.getScriptCache();

        //  Object with key/value pairs for all items found in cache.
        const data = shortCache.getAll(cacheKeys);
        const cachedDataList = [];

        for (const key of cacheKeys) {
            let parsedData = null;
            if (typeof data[key] !== 'undefined') {
                parsedData = JSON.parse(data[key]);
            }
            cachedDataList.push(parsedData);
        }

        return cachedDataList;
    }

    /**
     * Puts list of data into cache using one API call.  Data is converted to JSON before it is updated.
     * @param {String[]} cacheKeys 
     * @param {any[]} newCacheData 
     * @param {Number} cacheSeconds
     */
    static putFinanceValuesIntoShortCache(cacheKeys, newCacheData, cacheSeconds = 21600) {
        const bulkData = {};

        for (let i = 0; i < cacheKeys.length; i++) {
            if (CacheFinanceUtils.isValidGoogleValue(newCacheData[i])) {
                bulkData[cacheKeys[i]] = JSON.stringify(newCacheData[i]);
            }
        }

        const shortCache = CacheService.getScriptCache();
        shortCache.putAll(bulkData, cacheSeconds);
    }

    /**
     * Puts list of data into cache using one API call.  Data is converted to JSON before it is updated.
     * @param {String[]} cacheKeys 
     * @param {any[]} cacheData 
     * @param {Number} daysToHold
     */
    static putFinanceValuesIntoLongCache(cacheKeys, cacheData, daysToHold = 7) {
        const newCacheKeys = [];
        const newCacheData = [];

        for (let i = 0; i < cacheKeys.length; i++) {
            if (CacheFinanceUtils.isValidGoogleValue(cacheData[i])) {
                newCacheKeys.push(cacheKeys[i]);
                newCacheData.push(cacheData[i])
            }
        }

        ScriptSettings.putAllKeysWithData(newCacheKeys, newCacheData, daysToHold);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {String}
     */
    static makeCacheKey(symbol, attribute) {
        return `${attribute}|${symbol}`;
    }

    /**
     * It is common to have extra empty records loaded at end of table.
     * Remove those empty records at END of table only.
     * @param {any[][]} tableData 
     * @returns {any[][]}
     */
    static removeEmptyRecordsAtEndOfTable(tableData) {
        if (!Array.isArray(tableData)) {
            return tableData;
        }

        let blankLines = 0;
        for (let i = tableData.length - 1; i > 0; i--) {
            if (tableData[i].join().replace(/,/g, "").length > 0)
                break;
            blankLines++;
        }

        return tableData.slice(0, tableData.length - blankLines);
    }

    /**
     * 
     * @param {any} value 
     * @returns {Boolean}
     */
    static isValidGoogleValue(value) {
        return value !== null && typeof value !== 'undefined' && value !== "#N/A" && value !== '#ERROR!' && value !== '';
    }

    //  When you request a single column of data from getRange(), it is still a double array.
    //  Convert to single array for reguar array processing.
    static convertRowsToSingleArray(doubleArray) {
        if (! Array.isArray(doubleArray)) {
            return doubleArray;
        }
    
        return doubleArray.map(item => item[0]);
    }

    /**
     * 
     * @param {any[]} singleArray 
     * @returns {any[][]}
     */
    static convertSingleToDoubleArray(singleArray) {
        return singleArray.map(item => [item]);
    }
}