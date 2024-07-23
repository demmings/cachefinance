/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { ScriptSettings } from "./SQL/ScriptSettings.js";
import { ThirdPartyFinance, FinanceWebsiteSearch } from "./CacheFinance3rdParty.js";
import { cacheFinanceTest } from "./CacheFinanceTest.js";
import { StockAttributes, FinanceWebSites } from "./CacheFinanceWebSites.js";
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
    const val = CACHEFINANCE("TSE:CJP", "yieldpct");        // skipcq: JS-0128
    Logger.log(`Test CacheFinance FTN-A(yieldpct)=${val}`);
}

function testCacheFinances() {                                  // skipcq: JS-0128
    const symbols = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("A30:A165").getValues();
    const data = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("E30:E165").getValues();

    const cacheData = CACHEFINANCES(symbols, "PRICE", data);

    const singleSymbols = CacheFinanceUtils.convertRowsToSingleArray(symbols);

    Logger.log(`BULK CACHE TEST Success${cacheData} . ${singleSymbols}`);
}

/**
 * Enhancement to GOOGLEFINANCE function for stock/ETF symbols that a) return "#N/A" (temporary or consistently), b) data never available like 'yieldpct' for ETF's. 
 * @param {string} symbol - stock ticket with exchange (e.g.  "NYSEARCA:VOO")
 * @param {string} attribute - ["price", "yieldpct", "name"] - 
 * @param {any} googleFinanceValue - Optional.  Use GOOGLEFINANCE() to get default value, if '#N/A' will read cache.
 * BACKDOOR commands are entered using this parameter.
 *  "?" - List all backdoor abilities (SET, GET, SETBLOCKED, GETBLOCKED, LIST, REMOVE, CLEARCACHE, TEST)
 * e.g. =CACHEFINANCE("", "", "CLEARCACHE") or =CACHEFINANCE("TSE:CJP", "price", "GET")
 * @param {String} cmdOption - Option parameter used only with backdoor commands.
 * @returns {any}
 * @customfunction
 */
function CACHEFINANCE(symbol, attribute = "price", googleFinanceValue = "", cmdOption = "") {         // skipcq: JS-0128
    Logger.log(`CACHEFINANCE:${symbol}=${attribute}. Google=${googleFinanceValue}`);

    //  Special inputs that perform something other than a finance request.
    const providerUpdateMessage = CacheFinance.backDoorCommands(symbol, attribute, googleFinanceValue, cmdOption);
    if (providerUpdateMessage !== null) {
        return providerUpdateMessage;
    }

    if (symbol === '' || attribute === '') {
        return '';
    }

    const data = CACHEFINANCES([[symbol]], attribute, [[googleFinanceValue]]);

    return data[0][0];
}

/**
 * Bulk cache retrieval of finance data for updating large quantity of stock attributes.
 * @param {String[][]} symbols 
 * @param {String} attribute - ["price", "yieldpct", "name"]
 * @param {any[][]} defaultValues Default values from GoogleFinance()
 * @param {Number} webSiteLookupCacheSeconds Min. time between Web Lookups (max 21600 seconds)
 * @returns {any}
 * @customfunction
 */
function CACHEFINANCES(symbols, attribute = "price", defaultValues = [], webSiteLookupCacheSeconds = -1) {         // skipcq: JS-0128
    if (!Array.isArray(symbols)) {
        throw new Error("Expecting list of stock symbols.");
    }

    if (Array.isArray(symbols) && Array.isArray(defaultValues) && defaultValues.length > 0 && symbols.length !== defaultValues.length) {
        throw new Error("Stock symbol RANGE must match default values range.");
    }

    if (defaultValues === undefined || typeof defaultValues === 'string') {
        defaultValues = [];
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

    Logger.log(`CacheFinances START.  Attribute=${attribute} symbols=${symbols.length}`);

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

        //  GOOGLEFINANCE has failed OR was not used.  Is it in the cache?
        const shortCacheData = CacheFinance.getFinanceValueFromShortCache(cacheKey);
        if (shortCacheData !== null) {
            return shortCacheData;
        }

        //  Last resort... try other sites.
        const stockAttributes = ThirdPartyFinance.get(symbol, attribute);

        //  Failed third party lookup, try using long term cache.
        if (!stockAttributes.isAttributeSet(attribute)) {
            const longCacheData = CacheFinance.getFinanceValueFromLongCache(cacheKey);
            if (longCacheData !== null) {
                return longCacheData;
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
    static getBulkFinanceData(symbols, attribute, googleFinanceValues, webSiteLookupCacheSeconds = -1) {
        const MAX_SHORT_CACHE_SECONDS = 21600;      // For VALID GOOGLEFINANCE values.
        const MAX_SHORT_CACHE_THIRD_PARTY = 1200;   // This will force a lookup every 20 minutes for stocks NEVER found in GOOGLEFINANCE()

        //  ALL valid google data points are put in SHORT cache.
        CacheFinanceUtils.bulkShortCachePut(symbols, attribute, googleFinanceValues, MAX_SHORT_CACHE_SECONDS);

        //  All invalid data points with a valid entry in short cache is used.
        googleFinanceValues = CacheFinance.updateMissingValuesFromShortCache(symbols, attribute, googleFinanceValues);

        //  At this point, it will be mostly items that GOOGLE FINANCE just never works for.
        const symbolsWithNoData = CacheFinance.getSymbolsWithNoValidData(symbols, googleFinanceValues);

        //  Make requests (very slow) from financial web sites to find missing data.
        const thirdPartyStockAtributes = ThirdPartyFinance.getMissingStockAttributesFromThirdParty(symbolsWithNoData, attribute);
        const thirdPartyFinanceValues = CacheFinance.getValuesFromStockAttributes(thirdPartyStockAtributes, attribute);
        googleFinanceValues = CacheFinance.updateMasterWithMissed(symbols, googleFinanceValues, symbolsWithNoData, thirdPartyFinanceValues);

        //  All data found in websites (not GOOGLEFINANCE) is placed in cache (for a shorter period of time than those from GOOGLEFINANCE)
        const cacheSeconds = webSiteLookupCacheSeconds === -1 ? MAX_SHORT_CACHE_THIRD_PARTY : webSiteLookupCacheSeconds;
        CacheFinance.putAllStockAttributeDataIntoShortCache(thirdPartyStockAtributes, symbolsWithNoData, cacheSeconds);

        // Last, last resort.  Try to find in LONG CACHE.  This could be DAYS old, but it is better than invalid data.
        const lastResortMissingStocks = CacheFinance.getSymbolsWithNoValidData(symbols, googleFinanceValues);
        const longCacheValues = CacheFinanceUtils.bulkLongCacheGet(lastResortMissingStocks, attribute);
        googleFinanceValues = CacheFinance.updateMasterWithMissed(symbols, googleFinanceValues, lastResortMissingStocks, longCacheValues);

        //  Everything we need was in the short cache, so no need to update long cache.
        if (symbolsWithNoData.length > 0) {
            //  Save everything we have found into the long cache for dire use cases in future.
            CacheFinanceUtils.bulkLongCachePut(symbols, attribute, googleFinanceValues);
        }

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
        //  pulling from short cache is very slow, so if everything is GOOD it can be skipped.
        if (CacheFinance.isAllGoogleDefaultValuesValid(symbols, googleFinanceValues)) {
            return googleFinanceValues;
        }

        const valueFromCache = CacheFinanceUtils.bulkShortCacheGet(symbols, attribute).map(val => val === null ? "#N/A" : val);

        return googleFinanceValues.map((val, i) => CacheFinanceUtils.isValidGoogleValue(val) ? val : valueFromCache[i]);
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

        return CacheFinance.getSymbolsWithNoValidData(symbols, googleFinanceValues).length === 0;
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {any[]} googleFinanceValues 
     * @returns {String[]}
     */
    static getSymbolsWithNoValidData(symbols, googleFinanceValues) {
        return symbols.filter((_sym, i) => !CacheFinanceUtils.isValidGoogleValue(googleFinanceValues[i]));
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
     * @returns {any}
     */
    static getFinanceValueFromShortCache(cacheKey) {
        const shortCache = CacheService.getScriptCache();
        const data = shortCache.get(cacheKey);

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
        if (currentShortCacheValue === null)
            return true;

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

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @param {String} googleValue 
     * @param {String} cmdOption
     * @returns 
     */
    static backDoorCommands(symbol, attribute, googleValue, cmdOption) {
        const commandStr = googleValue.toString().toUpperCase().trim();
        cmdOption = cmdOption.toString().toUpperCase().trim();

        switch (commandStr) {
            case "":
                return null;

            case "?":
            case "HELP":
                return [["Valid commands in 3'rd parameter.  Erase after run to prevent future runs."],
                ["    ? (display help)"],
                ["    TEST (tests web sites)"],
                ["    CLEARCACHE (remove cache - run again if timeout. If symbol/attribute blank - removes all)"],
                ["    REMOVE (pref. site set as do not use site for symbol/attribute)"],
                ["    LIST (show all supported web lookups)"],
                ["    GET (current pref. site for symbol/attribute)"],
                ["    GETBLOCKED (current blocked site for symbol/attribute)"],
                ["    SET (4'th parm is set to pref. site for symbol/attribute)"],
                ["    SETBLOCKED (4'th parm is set to blocked site for symbol/attribute)"]];

            case "TEST":
                return cacheFinanceTest();

            case "CLEARCACHE":
                if (symbol !== "" && attribute !== "") {
                    CacheFinance.deleteFromCache(symbol, attribute);
                }
                else {
                    ScriptSettings.expire(true);
                }
                return 'Cache Cleared';

            case "REMOVE":
                return CacheFinance.removeCurrentProviderAsFavourite(symbol, attribute);

            case "GET":
                return CacheFinance.getCurrentProvider(symbol, attribute);

            case "GETBLOCKED":
                return CacheFinance.getBlockedProvider(symbol, attribute);

            case "SET":
                if (cmdOption !== "" && CacheFinance.listProviders().indexOf(cmdOption) === -1) {
                    return "Invalid provider name.  No change made.";
                }
                CacheFinance.setProviderAsFavourite(symbol, attribute, cmdOption);
                return `New provider (${cmdOption}) set as default for: ${symbol} ${attribute}`;

            case "SETBLOCKED":
                if (cmdOption !== "" && CacheFinance.listProviders().indexOf(cmdOption) === -1) {
                    return "Invalid provider name.  No change made.";
                }
                CacheFinance.setBlockedProvider(symbol, attribute, cmdOption);
                return `New provider (${cmdOption}) set as blocked for: ${symbol} ${attribute}`;

            case "LIST":
                return CacheFinanceUtils.convertSingleToDoubleArray(CacheFinance.listProviders());
        }

        return null;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {String}
     */
    static removeCurrentProviderAsFavourite(symbol, attribute) {
        CacheFinance.deleteFromCache(symbol, attribute);
        let statusMessage = "";

        const bestStockSites = FinanceWebsiteSearch.readBestStockWebsites();
        const objectKey = CacheFinanceUtils.makeCacheKey(symbol, attribute);
        Logger.log(`Removing current site for ${objectKey}`);

        if (typeof bestStockSites[objectKey] !== 'undefined') {
            const badSite = bestStockSites[objectKey];
            statusMessage = "Site removed for lookups: " + badSite;
            Logger.log(`Removing site from list: ${badSite}`);
            delete bestStockSites[objectKey];
            bestStockSites[CacheFinanceUtils.makeIgnoreSiteCacheKey(symbol, attribute)] = badSite;
            FinanceWebsiteSearch.writeBestStockWebsites(bestStockSites);
        }
        else {
            statusMessage = `Currently no preferred site for ${symbol} ${attribute}`;
        }

        return statusMessage;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @param {String} siteName 
     */
    static setProviderAsFavourite(symbol, attribute, siteName) {
        const objectKey = CacheFinanceUtils.makeCacheKey(symbol, attribute);
        CacheFinance.setProviderData(objectKey, siteName);
    }

    /**
     * Sets ONE web provider to NEVER be used for symbol/attribute.
     * @param {String} symbol 
     * @param {String} attribute 
     * @param {String} siteName 
     */
    static setBlockedProvider(symbol, attribute, siteName) {
        const objectKey = CacheFinanceUtils.makeIgnoreSiteCacheKey(symbol, attribute);
        CacheFinance.setProviderData(objectKey, siteName);
    }

    /**
     * 
     * @param {String} objectKey 
     * @param {String} siteName 
     */
    static setProviderData(objectKey, siteName) {
        const bestStockSites = FinanceWebsiteSearch.readBestStockWebsites();
        bestStockSites[objectKey] = siteName;

        if (siteName === "") {
            delete bestStockSites[objectKey];
        }

        FinanceWebsiteSearch.writeBestStockWebsites(bestStockSites);
    }

    /**
     * Returns the PREFERRED web site to find symbol/attribute.
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {String}
     */
    static getCurrentProvider(symbol, attribute) {
        const objectKey = CacheFinanceUtils.makeCacheKey(symbol, attribute);

        return CacheFinance.getProviderData(objectKey);
    }

    /**
     * Returns the web site provider that is never used to access symbol/attribute.
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {String}
     */
    static getBlockedProvider(symbol, attribute) {
        const objectKey = CacheFinanceUtils.makeIgnoreSiteCacheKey(symbol, attribute);

        return CacheFinance.getProviderData(objectKey);
    }

    /**
     * 
     * @param {String} objectKey 
     * @returns {String}
     */
    static getProviderData(objectKey) {
        let currentSite = "No site set.";
        const bestStockSites = FinanceWebsiteSearch.readBestStockWebsites();

        if (typeof bestStockSites[objectKey] !== 'undefined') {
            currentSite = bestStockSites[objectKey];
        }

        return currentSite;
    }

    /**
     * Returns all web site ID's used to retrieve stock info.
     * @returns {String[]}
     */
    static listProviders() {
        const webSites = new FinanceWebSites();

        const siteNames = webSites.siteList.map(site => site._siteName);

        return siteNames;
    }
}