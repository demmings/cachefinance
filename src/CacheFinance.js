/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { ScriptSettings } from "./SQL/ScriptSettings.js";
import { ThirdPartyFinance } from "./CacheFinance3rdParty.js";
import { cacheFinanceTest } from "./CacheFinanceTest.js";
import { StockAttributes } from "./CacheFinanceWebSites.js";
import { CacheService, SpreadsheetApp } from "./GasMocks.js";
import { CacheFinanceUtils } from "./CacheFinanceUtils.js";
export { CACHEFINANCE, CacheFinance, GOOGLEFINANCE_PARAM_NOT_USED };

class Logger {
    static log(msg) {
        console.log(msg);
    }
}
//  *** DEBUG END ***/

const GOOGLEFINANCE_PARAM_NOT_USED = "##NotSet##";

//  Function only used for testing in google sheets app script.
// skipcq: JS-0128
function testYieldPct() {
    const val = CACHEFINANCE("TSE:FTN-A", "yieldpct");        // skipcq: JS-0128
    Logger.log(`Test CacheFinance FTN-A(yieldpct)=${val}`);
}

function testCacheFinances() {
    // const symbols = [["ABC"], ["DEF"], ["GHI"], ["JKL"], ["TSE:FLJA"]];
    // const data = [[11.1], [22.2], [33.3], [44.4], ["#N/A"]];

    let symbols = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("A30:A165").getValues();
    const data = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("E30:E165").getValues();

    const cacheData = CACHEFINANCES(symbols, "PRICE", data);

    let singleSymbols = CacheFinanceUtils.convertRowsToSingleArray(symbols);

    Logger.log("BULK CACHE TEST Success");
}

/**
 * Enhancement to GOOGLEFINANCE function for stock/ETF symbols that a) return "#N/A" (temporary or consistently), b) data never available like 'yieldpct' for ETF's. 
 * @param {string} symbol 
 * @param {string} attribute - ["price", "yieldpct", "name"] - 
 * Special Attributes.
 *  "TEST" - returns test results from 3rd party sites.
 *  "CLEARCACHE" - Removes all Script Properties (used for long term cache) created by CACHEFINANCE
 * @param {any} googleFinanceValue - Optional.  Use GOOGLEFINANCE() to get value, if '#N/A' will read cache.
 * @returns {any}
 * @customfunction
 */
function CACHEFINANCE(symbol, attribute = "price", googleFinanceValue = GOOGLEFINANCE_PARAM_NOT_USED) {         // skipcq: JS-0128
    Logger.log(`CACHEFINANCE:${symbol}=${attribute}. Google=${googleFinanceValue}`);

    if (attribute.toUpperCase() === "TEST") {
        return cacheFinanceTest();
    }

    if (attribute.toUpperCase() === "CLEARCACHE") {
        ScriptSettings.expire(true);
        return 'Cache Cleared';
    }

    if (symbol === '' || attribute === '') {
        return '';
    }

    if (typeof googleFinanceValue === 'string' && googleFinanceValue === '') {
        googleFinanceValue = GOOGLEFINANCE_PARAM_NOT_USED;
    }

    return CacheFinance.getFinanceData(symbol, attribute, googleFinanceValue);
}

/**
 * Bulk cache retrieval of finance data for updating large quantity of stock attributes.
 * @param {String[][]} symbols 
 * @param {String} attribute - ["price", "yieldpct", "name"]
 * @param {any[][]} defaultValues Default values from GoogleFinance()
 * @param {Number} webSiteLookupCacheSeconds Min. time between Web Lookups (max 21600 seconds)
 * @returns 
 * @customfunction
 */
function CACHEFINANCES(symbols, attribute = "price", defaultValues = [], webSiteLookupCacheSeconds = -1) {         // skipcq: JS-0128
    if (!Array.isArray(symbols)) {
        throw new Error("Expecting list of stock symbols.");
    }

    if (Array.isArray(symbols) && Array.isArray(defaultValues) && defaultValues.length > 0 && symbols.length !== defaultValues.length) {
        throw new Error("Stock symbol RANGE must match default values range.");
    }

    const trimmedSymbols = CacheFinanceUtils.removeEmptyRecordsAtEndOfTable(symbols);
    const trimmedValues = CacheFinanceUtils.removeEmptyRecordsAtEndOfTable(defaultValues);

    //  Data ranges from sheets are double arrays.  Just make life simple and convert to single array.
    const newSymbols = CacheFinanceUtils.convertRowsToSingleArray(trimmedSymbols);
    const newValues = CacheFinanceUtils.convertRowsToSingleArray(trimmedValues);

    attribute = attribute.toUpperCase().trim();
    if (attribute === "CLEARCACHE") {
        ScriptSettings.expire(true);
        return 'Long Cache Cleared';
    }

    if (typeof newValues === 'string' && newValues.toUpperCase() === 'CLEARCACHE') {
        CacheFinanceUtils.bulkShortCacheRemoveAll(newSymbols, attribute);
        return 'Short Cache Cleared';
    }

    if (newSymbols.length === 0 || attribute === '') {
        return '';
    }

    Logger.log("CacheFinances START.  Attribute=" + attribute + " symbols=" + symbols.length);

    return CacheFinance.getBulkFinanceData(newSymbols, attribute, newValues, webSiteLookupCacheSeconds);
}

/**
 * @classdesc GOOGLEFINANCE helper function.  Returns default value (if available) and set this value to cache OR
 * reads from short term cache (<21600s) and returns value OR
 * reads from 3rd party screen scrapping OR
 * reads from long term cache
 */
class CacheFinance {
    /**
     * Replacement function to GOOGLEFINANCE for stock symbols not recognized by google.
     * @param {string} symbol 
     * @param {string} attribute - ["price", "yieldpct", "name"] 
     * @param {any} googleFinanceValue - Optional.  Use GOOGLEFINANCE() to get value, if '#N/A' will read cache.
     * @param {any} valueFromCache - optional - value previously read from cache.
     * @returns {any}
     */
    static getFinanceData(symbol, attribute, googleFinanceValue, valueFromCache = null) {
        attribute = attribute.toUpperCase().trim();
        symbol = symbol.toUpperCase();
        const cacheKey = CacheFinanceUtils.makeCacheKey(symbol, attribute);

        //  This time GOOGLEFINANCE worked!!!
        if (googleFinanceValue !== GOOGLEFINANCE_PARAM_NOT_USED && googleFinanceValue !== "#N/A" && googleFinanceValue !== '#ERROR!') {
            //  We cache here longer because we would normally be getting data from Google.
            //  If GoogleFinance is failing, we need the data to be held longer since it
            //  it is getting from cache as an emergency backup.
            CacheFinance.saveFinanceValueToCache(cacheKey, googleFinanceValue, 21600, valueFromCache);
            return googleFinanceValue;
        }

        //  In the case where GOOGLE never gives a value (or GoogleFinance is never used),
        //  we don't want to pull from long cache (at this point).
        const useShortCacheOnly = googleFinanceValue === GOOGLEFINANCE_PARAM_NOT_USED || googleFinanceValue !== "#N/A";

        //  GOOGLEFINANCE has failed OR was not used.  Is it in the cache?
        const data = CacheFinance.getFinanceValueFromCache(cacheKey, useShortCacheOnly);

        if (data !== null) {
            return data;
        }

        //  Last resort... try other sites.
        let stockAttributes = ThirdPartyFinance.get(symbol, attribute);

        //  Failed third party lookup, try using long term cache.
        if (!stockAttributes.isAttributeSet(attribute)) {
            const cachedStockAttribute = CacheFinance.getFinanceValueFromCache(cacheKey, false);
            if (cachedStockAttribute !== null) {
                stockAttributes = cachedStockAttribute;
            }
        }
        else {
            //  If we are mostly getting this finance item from a third party, we set the timeout
            //  a little shorter since we don't want to return extremely old data.
            CacheFinance.saveAllFinanceValuesToCache(symbol, stockAttributes);
        }

        return stockAttributes.getValue(attribute);
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {String} attribute 
     * @param {any[]} googleFinanceValues 
     * @param {Number} webSiteLookupCacheSeconds
     * @returns {any[][]}
     */
    static getBulkFinanceData(symbols, attribute, googleFinanceValues, webSiteLookupCacheSeconds) {
        const MAX_SHORT_CACHE_SECONDS = 21600;      // For VALID GOOGLEFINANCE values.
        const MAX_SHORT_CACHE_THIRD_PARTY = 1200;   // This will force a lookup every 20 minutes for stocks NEVER found in GOOGLEFINANCE()

        //  ALL valid google data points are put in SHORT cache.
        CacheFinanceUtils.bulkShortCachePut(symbols, attribute, googleFinanceValues, MAX_SHORT_CACHE_SECONDS);

        //  All invalid data points with a valid entry in short cache is used.
        googleFinanceValues = CacheFinance.updateMissingValuesFromShortCache(symbols, attribute, googleFinanceValues);

        //  At this point, it will be mostly items that GOOGLE FINANCE just never works for.
        let symbolsWithNoData = CacheFinance.getSymbolsWithNoValidData(symbols, googleFinanceValues);

        //  Make requests (very slow) from financial web sites to find missing data.
        const thirdPartyStockAtributes = ThirdPartyFinance.getMissingStockAttributesFromThirdParty(symbolsWithNoData, attribute);
        const thirdPartyFinanceValues = CacheFinance.getValuesFromStockAttributes(thirdPartyStockAtributes, attribute);
        googleFinanceValues = CacheFinance.updateMasterWithMissed(symbols, googleFinanceValues, symbolsWithNoData, thirdPartyFinanceValues);

        //  All data found in websites (not GOOGLEFINANCE) is placed in cache (for a shorter period of time than those from GOOGLEFINANCE)
        const cacheSeconds = webSiteLookupCacheSeconds === -1 ? MAX_SHORT_CACHE_THIRD_PARTY : webSiteLookupCacheSeconds;
        CacheFinance.putAllStockAttributeDataIntoShortCache(thirdPartyStockAtributes, symbolsWithNoData, cacheSeconds);

        // Last, last resort.  Try to find in LONG CACHE.  This could be DAYS old, but it is better than invalid data.
        symbolsWithNoData = CacheFinance.getSymbolsWithNoValidData(symbols, googleFinanceValues);
        const longCacheValues = CacheFinanceUtils.bulkLongCacheGet(symbolsWithNoData, attribute);
        googleFinanceValues = CacheFinance.updateMasterWithMissed(symbols, googleFinanceValues, symbolsWithNoData, longCacheValues);

        //  Save everything we have found into the long cache for dire use cases in future.
        CacheFinanceUtils.bulkLongCachePut(symbols, attribute, googleFinanceValues);

        return CacheFinanceUtils.convertSingleToDoubleArray(googleFinanceValues);
    }

    /**
     * 
     * @param {StockAttributes[]} stockAttributes 
     * @param {String} attribute 
     * @returns 
     */
    static getValuesFromStockAttributes(stockAttributes, attribute) {
        return stockAttributes.map(stockData => stockData.getValue(attribute));
    }

    /**
     * 
     * @param {StockAttributes[]} thirdPartyStockAtributes 
     * @param {String[]} symbolsWithNoData 
     * @param {Number} cacheSeconds 
     */
    static putAllStockAttributeDataIntoShortCache(thirdPartyStockAtributes, symbolsWithNoData, cacheSeconds) {
        const thirdPartyPriceValues = CacheFinance.getValuesFromStockAttributes(thirdPartyStockAtributes, "PRICE");
        const thirdPartyNameValues = CacheFinance.getValuesFromStockAttributes(thirdPartyStockAtributes, "NAME");
        const thirdPartyYieldValues = CacheFinance.getValuesFromStockAttributes(thirdPartyStockAtributes, "YIELDPCT");

        CacheFinanceUtils.bulkShortCachePut(symbolsWithNoData, "PRICE", thirdPartyPriceValues, cacheSeconds);
        CacheFinanceUtils.bulkShortCachePut(symbolsWithNoData, "NAME", thirdPartyNameValues, cacheSeconds);
        CacheFinanceUtils.bulkShortCachePut(symbolsWithNoData, "YIELDPCT", thirdPartyYieldValues, cacheSeconds);
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {String} attribute 
     * @param {any[]} googleFinanceValues 
     * @returns {any[]}
     */
    static updateMissingValuesFromShortCache(symbols, attribute, googleFinanceValues) {
        if (CacheFinance.isAllGoogleDefaultValuesValid(symbols, googleFinanceValues)) {
            return googleFinanceValues;
        }

        const updatedFinanceValues = [];
        const valueFromCache = CacheFinanceUtils.bulkShortCacheGet(symbols, attribute);
        for (let i = 0; i < symbols.length; i++) {
            if (CacheFinanceUtils.isValidGoogleValue(googleFinanceValues[i])) {
                updatedFinanceValues.push(googleFinanceValues[i]);
            } else {
                const valueToUseFromCache = valueFromCache[i] !== null ? valueFromCache[i] : "#N/A";
                updatedFinanceValues.push(valueToUseFromCache);
            }
        }

        return updatedFinanceValues;
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {any[]} googleFinanceValues 
     * @returns {Boolean}
     */
    static isAllGoogleDefaultValuesValid(symbols, googleFinanceValues) {
        if (symbols.length !== googleFinanceValues.length) {
            return false;
        }

        return  CacheFinance.getSymbolsWithNoValidData(symbols, googleFinanceValues).length === 0;
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {any[]} googleFinanceValues 
     * @returns {String[]}
     */
    static getSymbolsWithNoValidData(symbols, googleFinanceValues) {
        const symbolsWithNoData = [];

        for (let i = 0; i < symbols.length; i++) {
            if (!CacheFinanceUtils.isValidGoogleValue(googleFinanceValues[i])) {
                symbolsWithNoData.push(symbols[i]);
            }
        }

        return symbolsWithNoData;
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {any[]} googleFinanceValues 
     * @param {String[]} symbolsWithNoData 
     * @param {any[]} thirdPartyFinanceValues 
     * @returns 
     */
    static updateMasterWithMissed(symbols, googleFinanceValues, symbolsWithNoData, thirdPartyFinanceValues) {
        let startPos = 0;

        for (let i = 0; i < symbolsWithNoData.length; i++) {
            const j = symbols.indexOf(symbolsWithNoData[i], startPos);
            if (j !== -1) {
                if (CacheFinanceUtils.isValidGoogleValue(thirdPartyFinanceValues[i])) {
                    googleFinanceValues[j] = thirdPartyFinanceValues[i];
                }
                startPos = j;
            }
        }

        return googleFinanceValues;
    }

    /**
     * 
     * @param {String} cacheKey 
     * @param {Boolean} useShortCacheOnly
     * @returns {any} null if not found
     */
    static getFinanceValueFromCache(cacheKey, useShortCacheOnly) {
        const parsedData = CacheFinance.getFinanceValueFromShortCache(cacheKey);
        if (parsedData !== null || useShortCacheOnly) {
            return parsedData;
        }

        return CacheFinance.getFinanceValueFromLongCache(cacheKey);
    }

    /**
     * 
     * @param {String} cacheKey 
     * @returns {any}
     */
    static getFinanceValueFromShortCache(cacheKey) {
        const shortCache = CacheService.getScriptCache();

        const data = shortCache.get(cacheKey);

        //  Set to null while testing.  Remove when all is working.
        // data = null;

        if (data !== null && data !== "#ERROR!") {
            Logger.log(`Found in Short CACHE: ${cacheKey}. Value=${data}`);
            const parsedData = JSON.parse(data);
            if (!(typeof parsedData === 'string' && (parsedData === "#ERROR!" || parsedData === ""))) {
                return parsedData;
            }
        }

        return null;
    }

    /**
     * 
     * @param {String} cacheKey 
     * @returns {any}
     */
    static getFinanceValueFromLongCache(cacheKey) {
        const longCache = new ScriptSettings();

        const data = longCache.get(cacheKey);
        if (data !== null && data !== "#ERROR!") {
            Logger.log(`Long Term Cache.  Key=${cacheKey}. Value=${data}`);
            //  Long cache saves and returns same data type -so no conversion needed.
            return data;
        }

        return null;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {StockAttributes} stockAttributes 
     * @returns {void}
     */
    static saveAllFinanceValuesToCache(symbol, stockAttributes) {
        if (stockAttributes === null)
            return;
        if (stockAttributes.isAttributeSet("NAME"))
            CacheFinance.saveFinanceValueToCache(CacheFinanceUtils.makeCacheKey(symbol, "NAME"), stockAttributes.stockName, 1200);
        if (stockAttributes.isAttributeSet("PRICE"))
            CacheFinance.saveFinanceValueToCache(CacheFinanceUtils.makeCacheKey(symbol, "PRICE"), stockAttributes.stockPrice, 1200);
        if (stockAttributes.isAttributeSet("YIELDPCT"))
            CacheFinance.saveFinanceValueToCache(CacheFinanceUtils.makeCacheKey(symbol, "YIELDPCT"), stockAttributes.yieldPct, 1200);
    }

    /**
     * 
     * @param {String} key 
     * @param {any} financialData 
     * @param {Number} shortCacheSeconds 
     * @param {any} currentShortCacheValue
     * @returns {void}
     */
    static saveFinanceValueToCache(key, financialData, shortCacheSeconds = 1200, currentShortCacheValue = null) {
        const shortCache = CacheService.getScriptCache();
        if (currentShortCacheValue === null) {
            currentShortCacheValue = shortCache.get(key);
        }
        const longCacheDays = 7;

        if (!CacheFinance.isTimeToUpdateCache(currentShortCacheValue, financialData)) {
            return;
        }

        //  If we normally get the price from Google, we want to cache for a longer
        //  time because the only time we need a price for this particular stock
        //  is when GOOGLEFINANCE fails.
        let start = new Date().getTime();
        shortCache.put(key, JSON.stringify(financialData), shortCacheSeconds);
        const shortMs = new Date().getTime() - start;

        //  For emergency cases when GOOGLEFINANCE is down long term...
        start = new Date().getTime();
        const longCache = new ScriptSettings();
        longCache.put(key, financialData, longCacheDays);
        const longMs = new Date().getTime() - start;

        Logger.log(`SET GoogleFinance VALUE Long/Short Cache. Key=${key}.  Value=${financialData}. Short ms=${shortMs}. Long ms=${longMs}`);
    }

    /**
     * 
     * @param {String} currentShortCacheValue 
     * @param {any} financialData 
     * @returns {Boolean}
     */
    static isTimeToUpdateCache(currentShortCacheValue, financialData) {
        if (currentShortCacheValue !== null) {
            const oldData = JSON.parse(currentShortCacheValue);

            if (oldData === financialData) {
                Logger.log("GoogleFinance VALUE.  No Change in SHORT Cache.");
                return false;
            }

            if (oldData > 0 && financialData > 0) {
                const changeInPrice = oldData - financialData;
                const percentChange = Math.abs(changeInPrice / oldData);
                if (percentChange < 0.0025) {
                    Logger.log(`Short Cache Changed very little.  Old=${oldData} . New=${financialData}`);
                    return false;
                }
            }

            Logger.log(`Short Cache Changed.  Old=${oldData} . New=${financialData}`);
        }

        return true;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     */
    static deleteFromCache(symbol, attribute) {
        const key = CacheFinanceUtils.makeCacheKey(symbol, attribute);

        CacheFinance.deleteFromShortCache(key);
        CacheFinance.deleteFromLongCache(key);
    }

    /**
     * 
     * @param {String} key 
     */
    static deleteFromLongCache(key) {
        const longCache = new ScriptSettings();

        const currentLongCacheValue = longCache.get(key);
        if (currentLongCacheValue !== null) {
            longCache.delete(key);
        }
    }

    /**
     * 
     * @param {String} key 
     */
    static deleteFromShortCache(key) {
        const shortCache = CacheService.getScriptCache();

        const currentShortCacheValue = shortCache.get(key);
        if (currentShortCacheValue !== null) {
            shortCache.remove(key);
        }
    }
}