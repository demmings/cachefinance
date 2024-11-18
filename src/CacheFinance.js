/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { ScriptSettings } from "./SQL/ScriptSettings.js";
import { ThirdPartyFinance, FinanceWebsiteSearch } from "./CacheFinance3rdParty.js";
import { cacheFinanceTest } from "./CacheFinanceTest.js";
import { StockAttributes, FinanceWebSites } from "./CacheFinanceWebSites.js";
import { CacheService, SpreadsheetApp } from "./GasMocks.js";
import { CacheFinanceUtils } from "./CacheFinanceUtils.js";
export { CACHEFINANCE, CacheFinance };

class Logger {
    static log(msg) {
        console.log(msg);
    }
}

//  Function only used for testing in google sheets app script.
// skipcq: JS-0128
function testYieldPct() {
    const val = CACHEFINANCE("TSE:CJP", "yieldpct");        // skipcq: JS-0128
    Logger.log(`Test CacheFinance TSE:CJP(yieldpct)=${val}`);
}

function testCacheFinances() {                                  // skipcq: JS-0128
    const symbols = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("A30:A165").getValues();
    const data = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("E30:E165").getValues();

    const cacheData = CACHEFINANCES(symbols, "PRICE", data);

    const singleSymbols = CacheFinanceUtils.convertRowsToSingleArray(symbols);

    Logger.log(`BULK CACHE TEST Success${cacheData} . ${singleSymbols}`);
}

function testUpdateMaster() {
    const symbols = ["TSE:ZTL", "TSE:FTN-A", "TSE:ZTL"];
    const googleFinanceValues = [null, 10.0, null];
    const symbolsWithNoData = ["TSE:ZTL"];
    const thirdPartyFinanceValues = [15];

    const newGoogleFinance = CacheFinance.updateMasterWithMissed(symbols, googleFinanceValues, symbolsWithNoData, thirdPartyFinanceValues);
    Logger.log(newGoogleFinance);
}
//  *** DEBUG END ***/

/**
 * Enhancement to GOOGLEFINANCE function for stock/ETF symbols that a) return "#N/A" (temporary or consistently), b) data never available like 'yieldpct' for ETF's. 
 * @param {string} symbol - stock ticket with exchange (e.g.  "NYSEARCA:VOO")
 * @param {string} attribute - ["price", "yieldpct", "name"] - 
 * @param {any} googleFinanceValue - Optional.  Use GOOGLEFINANCE() to get default value, if '#N/A' will read cache.
 * BACKDOOR commands are entered using this parameter.
 *  "?" - List all backdoor abilities (SET, GET, SETBLOCKED, GETBLOCKED, LIST, REMOVE, CLEARCACHE, EXPIRECACHE, TEST)
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
    let isSingleLookup = false;
    if (!Array.isArray(symbols) && !Array.isArray(defaultValues)) {
        isSingleLookup = true;
        symbols = [[symbols]];
        defaultValues = [[defaultValues]];
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
    const singleSymbols = CacheFinanceUtils.convertRowsToSingleArray(trimmedSymbols);
    const newValues = CacheFinanceUtils.convertRowsToSingleArray(trimmedValues);

    const newSymbols = singleSymbols.map(sym => sym.toUpperCase());
    attribute = attribute.toUpperCase().trim();

    if (newSymbols.length === 0 || attribute === '') {
        return '';
    }

    Logger.log(`CacheFinances START.  Attribute=${attribute} symbols=${symbols.length} websiteLookupSeconds=${webSiteLookupCacheSeconds}`);

    let financeValues = CacheFinance.getBulkFinanceData(newSymbols, attribute, newValues, webSiteLookupCacheSeconds);
    if (isSingleLookup) {
        financeValues = financeValues[0][0];
    }

    return financeValues;
}

/**
 * @classdesc GOOGLEFINANCE helper function.  Returns default value (if available) and set this value to cache OR
 * reads from short term cache (<21600s) and returns value OR
 * reads from 3rd party screen scrapping OR
 * reads from long term cache
 */
class CacheFinance {
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
        //  All data found in websites (not GOOGLEFINANCE) is placed in cache (for a shorter period of time than those from GOOGLEFINANCE)
        const cacheSeconds = webSiteLookupCacheSeconds === -1 ? MAX_SHORT_CACHE_THIRD_PARTY : webSiteLookupCacheSeconds;
        CacheFinanceUtils.bulkShortCachePut(symbolsWithNoData, attribute, thirdPartyFinanceValues, cacheSeconds);

        googleFinanceValues = CacheFinance.updateMasterWithMissed(symbols, googleFinanceValues, symbolsWithNoData, thirdPartyFinanceValues);

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
        const updatedValues = [];
        for (let i = 0; i < symbols.length; i++) {
            const val = CacheFinanceUtils.isValidGoogleValue(googleFinanceValues[i]) ? googleFinanceValues[i] : valueFromCache[i];
            updatedValues.push(val);
        }

        return updatedValues;
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
        const noInfoSymbols = symbols.filter((_sym, i) => !CacheFinanceUtils.isValidGoogleValue(googleFinanceValues[i]));

        // @ts-ignore
        return [... new Set(noInfoSymbols)];
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {any[]} googleFinanceValues 
     * @param {String[]} symbolsWithNoData 
     * @param {any[]} thirdPartyFinanceValues 
     * @returns {any[]}
     */
    static updateMasterWithMissed(symbols, googleFinanceValues, symbolsWithNoData, thirdPartyFinanceValues) {
        for (let i = 0; i < symbolsWithNoData.length; i++) {
            let startPos = 0;

            while (startPos !== -1) {
                startPos = symbols.indexOf(symbolsWithNoData[i], startPos);
                if (startPos !== -1) {
                    if (CacheFinanceUtils.isValidGoogleValue(thirdPartyFinanceValues[i])) {
                        googleFinanceValues[startPos] = thirdPartyFinanceValues[i];
                    }
                    startPos++;
                }
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
                ["    EXPIRECACHE (removes OLD cached items)"],
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

            case "EXPIRECACHE":
                ScriptSettings.expire(false);
                return 'Old Cache Items Removed';

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

            default:
                return null;
        }
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
            statusMessage = `Site removed for lookups: ${badSite}`;
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