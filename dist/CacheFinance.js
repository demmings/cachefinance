
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
                if (cmdOption !== "" && !CacheFinance.listProviders().includes(cmdOption)) {
                    return "Invalid provider name.  No change made.";
                }
                CacheFinance.setProviderAsFavourite(symbol, attribute, cmdOption);
                return `New provider (${cmdOption}) set as default for: ${symbol} ${attribute}`;

            case "SETBLOCKED":
                if (cmdOption !== "" && !CacheFinance.listProviders().includes(cmdOption)) {
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

        if (bestStockSites[objectKey] === undefined) {
            statusMessage = `Currently no preferred site for ${symbol} ${attribute}`;
        }
        else {
            const badSite = bestStockSites[objectKey];
            statusMessage = `Site removed for lookups: ${badSite}`;
            Logger.log(`Removing site from list: ${badSite}`);
            delete bestStockSites[objectKey];
            bestStockSites[CacheFinanceUtils.makeIgnoreSiteCacheKey(symbol, attribute)] = badSite;
            FinanceWebsiteSearch.writeBestStockWebsites(bestStockSites);
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

        if (bestStockSites[objectKey] !== undefined) {
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

/** @classdesc 
 * Stores settings for the SCRIPT.  Long term cache storage for small tables.  */
class ScriptSettings {      //  skipcq: JS-0128
    /**
     * For storing cache data for very long periods of time.
     */
    constructor() {
        this.scriptProperties = PropertiesService.getScriptProperties();
    }

    /**
     * Get script property using key.  If not found, returns null.
     * @param {String} propertyKey 
     * @returns {any}
     */
    get(propertyKey) {
        const myData = this.scriptProperties.getProperty(propertyKey);

        if (myData === null)
            return null;

        /** @type {PropertyData} */
        const myPropertyData = JSON.parse(myData);

        if (PropertyData.isExpired(myPropertyData))
        {
            this.delete(propertyKey);
            return null;
        }

        return PropertyData.getData(myPropertyData);
    }

    /**
     * Put data into our PROPERTY cache, which can be held for long periods of time.
     * @param {String} propertyKey - key to finding property data.
     * @param {any} propertyData - value.  Any object can be saved..
     * @param {Number} daysToHold - number of days to hold before item is expired.
     */
    put(propertyKey, propertyData, daysToHold = 1) {
        //  Create our object with an expiry time.
        const objData = new PropertyData(propertyData, daysToHold);

        //  Our property needs to be a string
        const jsonData = JSON.stringify(objData);

        try {
            this.scriptProperties.setProperty(propertyKey, jsonData);
        }
        catch (ex) {
            throw new Error("Cache Limit Exceeded.  Long cache times have limited storage available.  Only cache small tables for long periods.");
        }
    }

    /**
     * @param {Object} propertyDataObject 
     * @param {Number} daysToHold 
     */
    putAll(propertyDataObject, daysToHold = 1) {
        const keys = Object.keys(propertyDataObject);
        for (const key of keys) {
            this.put(key, propertyDataObject[key], daysToHold);
        }
    }

    /**
     * Puts list of data into cache using one API call.  Data is converted to JSON before it is updated.
     * @param {String[]} cacheKeys 
     * @param {any[]} newCacheData 
     * @param {Number} daysToHold
     */
    static putAllKeysWithData(cacheKeys, newCacheData, daysToHold = 7) {
        const bulkData = {};

        for (let i = 0; i < cacheKeys.length; i++) {
            //  Create our object with an expiry time.
            const objData = new PropertyData(newCacheData[i], daysToHold);

            //  Our property needs to be a string
            bulkData[cacheKeys[i]] = JSON.stringify(objData);
        }

        PropertiesService.getScriptProperties().setProperties(bulkData);
    }

    /**
     * Returns ALL cached data for each key value requested. 
     * Only 1 API call is made, so much faster than retrieving single values.
     * @param {String[]} cacheKeys 
     * @returns {any[]}
     */
    static getAll(cacheKeys) {
        const values = [];

        if (cacheKeys.length === 0) {
            return values;
        }
        
        const allProperties = PropertiesService.getScriptProperties().getProperties();

        //  Removing properties is very slow, so remove only 1 at a time.  This is enough as this function is called frequently.
        ScriptSettings.expire(false, 1, allProperties);

        for (const key of cacheKeys) {
            const myData = allProperties[key];

            if (myData === undefined) {
                values.push(null);
            }
            else {
                /** @type {PropertyData} */
                const myPropertyData = JSON.parse(myData);

                if (PropertyData.isExpired(myPropertyData)) {
                    values.push(null);
                    PropertiesService.getScriptProperties().deleteProperty(key);
                    Logger.log(`Delete expired Script Property Key=${key}`);
                }
                else {
                    values.push(PropertyData.getData(myPropertyData));
                }
            }
        }

        return values;
    }

    /**
     * Removes script settings that have expired.
     * @param {Boolean} deleteAll - true - removes ALL script settings regardless of expiry time.
     * @param {Number} maxDelete - maximum number of items to delete that are expired.
     * @param {Object} allPropertiesObject - All properties already loaded.  If null, will load iteself.
     */
    static expire(deleteAll, maxDelete = 999, allPropertiesObject = null) {
        const allProperties = allPropertiesObject === null ? PropertiesService.getScriptProperties().getProperties() : allPropertiesObject;
        const allKeys = Object.keys(allProperties);
        let deleteCount = 0;

        for (const key of allKeys) {
            let propertyValue = null;
            try {
                propertyValue = JSON.parse(allProperties[key]);
            }
            catch (e) {
                //  A property that is NOT cached by CACHEFINANCE
                continue;
            }

            const propertyOfThisApplication = propertyValue?.expiry !== undefined;

            if (propertyOfThisApplication && (PropertyData.isExpired(propertyValue) || deleteAll)) {
                PropertiesService.getScriptProperties().deleteProperty(key);
                delete allProperties[key];

                //  There is no way to iterate existing from 'short' cache, so we assume there is a
                //  matching short cache entry and attempt to delete.
                CacheFinance.deleteFromShortCache(key);

                Logger.log(`Removing expired SCRIPT PROPERTY: key=${key}`);

                deleteCount++;
            }

            if (deleteCount >= maxDelete) {
                return;
            }
        }
    }

    /**
     * Delete a specific key in script properties.
     * @param {String} key 
     */
    delete(key) {
        if (this.scriptProperties.getProperty(key) !== null) {
            this.scriptProperties.deleteProperty(key);
        }
    }
}

/**
 * @classdesc Converts data into JSON for getting/setting in ScriptSettings.
 */
class PropertyData {
    /**
     * 
     * @param {any} propertyData 
     * @param {Number} daysToHold 
     */
    constructor(propertyData, daysToHold) {
        const someDate = new Date();

        /** @property {String} */
        this.myData = JSON.stringify(propertyData);
        /** @property {Date} */
        this.expiry = someDate.setMinutes(someDate.getMinutes() + daysToHold * 1440);
    }

    /**
     * @param {PropertyData} obj 
     * @returns {any}
     */
    static getData(obj) {
        let value = null;
        try {
            value = JSON.parse(obj.myData);
        }
        catch (ex) {
            Logger.log(`Invalid property value.  Not JSON: ${ex.toString()}`);
        }

        return value;
    }

    /**
     * 
     * @param {PropertyData} obj 
     * @returns {Boolean}
     */
    static isExpired(obj) {
        const someDate = new Date();
        const expiryDate = new Date(obj.expiry);
        return (expiryDate.getTime() < someDate.getTime())
    }
}

/**
 * @classdesc Find STOCK/ETF data by scraping data from 3rd party finance websites.
 */
class ThirdPartyFinance {                   //  skipcq: JS-0128
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    static get(symbol, attribute) {
        const searcher = new FinanceWebsiteSearch();
        const data = searcher.get(symbol, attribute);

        return data;
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {String} attribute 
     * @returns {StockAttributes[]}
     */
    static getMissingStockAttributesFromThirdParty(symbols, attribute) {
        const data = FinanceWebsiteSearch.getAll(symbols, attribute);

        return data;
    }
}

/**
 * @classdesc Make a plan to find finance data from websites and then execute the plan and return the data.
 */
class FinanceWebsiteSearch {
    constructor() {
        const siteInfo = new FinanceWebSites();

        /** @type {FinanceWebSite[]} */
        this.financeSiteList = siteInfo.get();
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    get(symbol, attribute) {
        const dataPlan = this.getLookupPlan(symbol, attribute);
        if (dataPlan.lookupPlan === null) {
            return new StockAttributes();
        }

        return dataPlan.data;
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {String} attribute 
     * @returns {StockAttributes[]}
     */
    static getAll(symbols, attribute) {
        if (symbols.length === 0) {
            return [];
        }

        const MAX_TIME_FOR_FETCH_Ms = 25000;        //  Custom function times out at 30 seconds, so we need to limit.
        const bestStockSites = FinanceWebsiteSearch.readBestStockWebsites();
        const siteURLs = FinanceWebsiteSearch.getAllStockWebSiteFunctions(symbols, attribute, bestStockSites);

        let batch = 1;
        const startTime = Date.now();
        let missingStockData = [...siteURLs];
        while (missingStockData.length > 0 && (Date.now() - startTime) < MAX_TIME_FOR_FETCH_Ms) {
            const [URLs, batchUsedStockSites] = FinanceWebsiteSearch.getNextUrlBatch(missingStockData);

            Logger.log(`Batch=${batch}. URLs. ${URLs}`);
            const responses = FinanceWebsiteSearch.bulkSiteFetch(URLs);
            Logger.log(`Batch=${batch}. Responses=${responses.length}. Total Elapsed=${Date.now() - startTime}`);
            batch++;

            FinanceWebsiteSearch.updateStockResults(batchUsedStockSites, URLs, responses, attribute, bestStockSites);

            missingStockData = missingStockData.filter(stock => !stock.stockAttributes.isAttributeSet(attribute) && !stock.isSitesDone())
        }

        //  Note:  If separate CACHEFINANCES() run at the same time, the last process to finish will overwrite any new results
        //         from the other runs.  This is not critical, since it is ONLY used to improve the ordering of sites to call AND
        //         over time as the processes run on their own, the data will be corrected.
        FinanceWebsiteSearch.writeBestStockWebsites(bestStockSites);

        return siteURLs.map(stock => stock.stockAttributes);
    }

    /**
     * Create a batch of URL's that will lookup our stock info in one large fetch.
     * @param {StockWebURL[]} missingStockData 
     * @returns {[String[], StockWebURL[]]}
     */
    static getNextUrlBatch(missingStockData) {
        const MAX_FETCHALL_BATCH_SIZE = 50;
        const URLs = [];
        const batchUsedStockSites = [];

        for (const stockData of missingStockData) {
            let URL = "";
            while (URL === "" && !stockData.isSitesDone()) {
                if (FinanceWebsiteSearch.canRequestNow(stockData)) {
                    URL = stockData.getURL();
                }
                else {
                    stockData.skipToNextSite();
                }
            }

            if (URL !== "") {
                URLs.push(URL);
                batchUsedStockSites.push(stockData);
            }

            if (URLs.length >= MAX_FETCHALL_BATCH_SIZE) {
                break;
            }
        }

        //  Update throttle status back to CACHE.
        //  Updating CACHE is very slow, so it is done ONCE for each site after getting URL's.
        for (const stockData of missingStockData) {
            const throttleObject = stockData.getThrottleObject();
            throttleObject?.update();
        }

        return [URLs, batchUsedStockSites];
    }

    /**
     * 
     * @param {StockWebURL} stockData 
     * @returns {Boolean}
     */
    static canRequestNow(stockData) {
        const URL = stockData.getURL();
        if (URL === null || URL === '') {
            return false;
        }

        const throttleObject = stockData.getThrottleObject();
        return !(throttleObject !== null && !throttleObject.checkAndIncrement());
    }

    /**
     * 
     * @param {StockWebURL[]} missingStockData 
     * @param {String[]} URLs 
     * @param {String[]} responses 
     * @param {String} attribute
     * @param {Object} bestStockSites
     */
    static updateStockResults(missingStockData, URLs, responses, attribute, bestStockSites) {
        for (let i = 0; i < URLs.length; i++) {
            missingStockData[i].parseResponse(responses[i], attribute);
            missingStockData[i].updateBestSites(bestStockSites, attribute);
            missingStockData[i].skipToNextSite();
        }
    }

    /**
     * 
     * @param {String[]} symbols 
     * @param {String} attribute 
     * @param {Object} bestStockSites
     * @returns {StockWebURL[]}
     */
    static getAllStockWebSiteFunctions(symbols, attribute, bestStockSites) {
        const stockURLs = [];
        const apiMap = new Map();
        const throttleMap = new Map();
        const siteInfo = new FinanceWebSites();
        const siteList = siteInfo.get();

        //  Getting this is slow, so save and use later.
        for (const site of siteList) {
            apiMap.set(site.siteName, site.siteObject.getApiKey())
            throttleMap.set(site.siteName, site.siteObject.getThrottleObject())
        }

        for (const symbol of symbols) {
            const stockURL = new StockWebURL(symbol);
            const bestSite = bestStockSites[CacheFinanceUtils.makeCacheKey(symbol, attribute)];
            const skipSite = bestStockSites[CacheFinanceUtils.makeIgnoreSiteCacheKey(symbol, attribute)];

            for (const site of siteList) {
                stockURL.addSiteURL(site.siteName, bestSite, skipSite, site.siteObject.getURL(symbol, attribute, apiMap.get(site.siteName)), site.siteObject.parseResponse, throttleMap.get(site.siteName));
            }

            stockURLs.push(stockURL);
        }

        return stockURLs;
    }

    /**
     * 
     * @returns {Object}
     */
    static readBestStockWebsites() {
        const longCache = new ScriptSettings();
        const siteObject = longCache.get("CACHE_WEBSITES");

        return siteObject === null ? {} : siteObject;
    }

    /**
     * 
     * @param {Object} siteObject 
     */
    static writeBestStockWebsites(siteObject) {
        const longCache = new ScriptSettings();
        longCache.put("CACHE_WEBSITES", siteObject, 1825);
    }


    /**
     * 
     * @param {String[]} URLs 
     * @returns {String[]}
     */
    static bulkSiteFetch(URLs) {
        const filteredURLs = URLs.filter(url => url.trim() !== '');
        const fetchURLs = filteredURLs.map(url => {
            return {
                'url': url,         // skipcq: JS-0240 
                'method': 'get',
                'muteHttpExceptions': true
            }
        });

        let dataSet = [];
        const rawSiteData = UrlFetchApp.fetchAll(fetchURLs);
        dataSet = rawSiteData.map(response => response.getContentText());

        return dataSet;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static makeCacheKey(symbol) {
        return `WebSearch|${symbol}`;
    }

    /**
     * @typedef {Object} PlanPlusData
     * @property {FinanceSiteLookupAnalyzer} lookupPlan
     * @property {StockAttributes} data
     */

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {PlanPlusData}
     */
    getLookupPlan(symbol, attribute) {
        let data = null;
        const LOOKUP_PLAN_ACTIVE_DAYS = 7;
        const longCache = new ScriptSettings();

        const cacheKey = FinanceWebsiteSearch.makeCacheKey(symbol);
        const lookupPlan = longCache.get(cacheKey);

        if (lookupPlan === null) {
            const planWithData = this.createLookupPlan(symbol);

            // Create small object with needed data for conversion to JSON.
            const searchPlan = planWithData.lookupPlan.createFinanceSiteList();

            longCache.put(cacheKey, searchPlan, LOOKUP_PLAN_ACTIVE_DAYS);
            return planWithData;
        }
        else {
            //  Find stock attributes using optimal search plan.
            data = FinanceSiteLookupAnalyzer.getStockAttribute(lookupPlan, attribute);
        }

        return { lookupPlan, data };
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {PlanPlusData}
     */
    createLookupPlan(symbol) {
        const plans = [];

        for (const site of this.financeSiteList) {
            const sitePlan = new FinanceSiteLookupStats(symbol, site);

            const startTime = Date.now();
            const data = site.siteObject.getInfo(symbol);

            sitePlan.setSearchTime(Date.now() - startTime)
                .setAttributes(data);

            plans.push(sitePlan);
        }

        const lookupPlan = new FinanceSiteLookupAnalyzer(symbol);
        lookupPlan.analyzeSiteStatus(plans);
        const data = lookupPlan.getStockAttributes();

        return { lookupPlan, data };
    }

    /**
     * Delete a stock lookup plan.
     * @param {String} symbol 
     */
    static deleteLookupPlan(symbol) {
        const longCache = new ScriptSettings();

        const cacheKey = FinanceWebsiteSearch.makeCacheKey(symbol);

        longCache.delete(cacheKey);
    }
}

/**
 * @classdesc Tracks and generates URLs to find stock data.  Also tracks parsing functions for extracting data from HTML.
 */
class StockWebURL {
    constructor(symbol) {
        this.symbol = symbol;
        this.siteName = [];
        this.siteURL = [];
        this.bestSites = [];
        this.parseFunction = [];
        this.throttleObject = [];
        /** @type {StockAttributes} */
        this.stockAttributes = new StockAttributes();
        this.siteIterator = 0;
    }

    /**
     * 
     * @param {String} siteName 
     * @param {String} URL 
     * @param {Object} parseResponseFunction 
     * @param {SiteThrottle} throttleObject
     * @returns 
     */
    addSiteURL(siteName, bestSite, skipSite, URL, parseResponseFunction, throttleObject) {
        if (URL.trim() === '' || siteName === skipSite) {
            return;
        }

        if (siteName === bestSite) {
            this.siteName.unshift(siteName);
            this.siteURL.unshift(URL);
            this.parseFunction.unshift(parseResponseFunction);
            this.throttleObject.unshift(throttleObject);
            this.bestSites.unshift(true);
        }
        else {
            this.siteName.push(siteName);
            this.siteURL.push(URL);
            this.parseFunction.push(parseResponseFunction);
            this.throttleObject.push(throttleObject);
            this.bestSites.push(false);
        }
    }

    /**
     * Returns next website URL to be used.
     * @returns {String}
     */
    getURL() {
        return this.siteIterator < this.siteURL.length ? this.siteURL[this.siteIterator] : null;
    }

    /**
     * 
     * @returns {SiteThrottle}
     */
    getThrottleObject() {
        return this.siteIterator < this.throttleObject.length ? this.throttleObject[this.siteIterator] : null;
    }

    /**
     * 
     * @param {String} html 
     * @param {String} attribute ["PRICE, "NAME, "YIELDPCT"]
     * @returns {StockAttributes}
     */
    parseResponse(html, attribute) {
        this.stockAttributes = this.siteIterator < this.siteURL.length ? this.parseFunction[this.siteIterator](html, this.symbol, attribute) : null;

        //  Keep track of a website that worked, so we use right away next time.
        this.bestSites[this.siteIterator] = this?.stockAttributes.isAttributeSet(attribute);

        return this.stockAttributes;
    }

    /**
     * 
     * @param {Object} bestStockSites 
     * @param {String} attribute 
     */
    updateBestSites(bestStockSites, attribute) {
        const key = CacheFinanceUtils.makeCacheKey(this.symbol, attribute);
        bestStockSites[key] = (this?.stockAttributes.isAttributeSet(attribute)) ?  this.siteName[this.siteIterator] : "";
    }

    /**
     * Updates internal pointer for next site to be used.
     * @returns {void}
     */
    skipToNextSite() {
        this.siteIterator++;
    }

    /**
     * 
     * @returns {Boolean}
     */
    isSitesDone() {
        return this.siteIterator >= this.siteName.length;
    }
}

/**
 * @classdesc Ordered list of sites for doing attribute lookup.
 * This object will be converted to JSON for storage.
 */
class FinanceSiteList {
    /**
     * Initialize object to store optimal lookup sites for given stock symbol.
     * @param {String} stockSymbol 
     */
    constructor(stockSymbol) {
        this.symbol = stockSymbol;
        /** @property {String[]} */
        this.priceSites = [];
        /** @property {String[]} */
        this.nameSites = [];
        /** @property {String[]} */
        this.yieldSites = [];
    }

    /**
     * 
     * @param {String[]} arr 
     * @returns {FinanceSiteList}
     */
    setPriceSites(arr) {
        this.priceSites = [...arr];

        return this;
    }

    /**
     * 
     * @param {String[]} arr 
     * @returns {FinanceSiteList}
     */
    setNameSites(arr) {
        this.nameSites = [...arr];

        return this;
    }

    /**
     * 
     * @param {String[]} arr 
     * @returns {FinanceSiteList}
     */
    setYieldSites(arr) {
        this.yieldSites = [...arr];

        return this;
    }
}

/**
 * @classdesc For analyzing finance websites.
 */
class FinanceSiteLookupAnalyzer {
    /**
     * Initialize object to compare finance sites for completeness and speed.
     * @param {String} symbol 
     */
    constructor(symbol) {
        this.symbol = symbol;
        /** @property {FinanceSiteLookupStats[]} */
        this.priceSites = [];
        /** @property {FinanceSiteLookupStats[]} */
        this.nameSites = [];
        /** @property {FinanceSiteLookupStats[]} */
        this.yieldSites = [];
    }

    /**
     * 
     * @returns {FinanceSiteList}
     */
    createFinanceSiteList() {
        const orderedFinances = new FinanceSiteList(this.symbol);

        orderedFinances.setPriceSites(this.priceSites.map(a => a.siteObject.siteName))
            .setNameSites(this.nameSites.map(a => a.siteObject.siteName))
            .setYieldSites(this.yieldSites.map(a => a.siteObject.siteName));

        return orderedFinances;
    }

    /**
     * 
     * @param {FinanceSiteLookupStats[]} siteStats 
     */
    analyzeSiteStatus(siteStats) {
        this.priceSites = siteStats.filter(a => a.price !== null);
        this.nameSites = siteStats.filter(a => a.name !== null);
        this.yieldSites = siteStats.filter(a => a.yield !== null);

        this.priceSites.sort((a, b) => a.timeMs - b.timeMs);
        this.nameSites.sort((a, b) => a.timeMs - b.timeMs);
        this.yieldSites.sort((a, b) => a.timeMs - b.timeMs);
    }


    /**
     * Returns most recent stock attributes found when analyzing sites.
     * @returns 
     */
    getStockAttributes() {
        const attributes = new StockAttributes();

        attributes.stockPrice = this.priceSites.length > 0 ? this.priceSites[0].price : null;
        attributes.stockName = this.nameSites.length > 0 ? this.nameSites[0].name : null;
        attributes.yieldPct = this.yieldSites.length > 0 ? this.yieldSites[0].yield : null;

        return attributes;
    }

    /**
     * 
     * @param {FinanceSiteList} stockSites 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    static getStockAttribute(stockSites, attribute) {
        let siteArr = [];

        switch (attribute) {
            case "PRICE":
                siteArr = stockSites.priceSites;
                break;

            case "YIELDPCT":
                siteArr = stockSites.yieldSites;
                break;

            case "NAME":
                siteArr = stockSites.nameSites;
                break;

            default:
                siteArr = stockSites.priceSites;
                break;
        }

        return FinanceSiteLookupAnalyzer.getAttributeDataFromSite(stockSites.symbol, siteArr, attribute);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String[]} siteArr 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    static getAttributeDataFromSite(symbol, siteArr, attribute) {
        let data = new StockAttributes();

        const sitesSearchFunction = new FinanceWebSites();
        for (const site of siteArr) {
            const siteFunction = sitesSearchFunction.getByName(site);
            if (siteFunction === null) {
                Logger.log(`Invalid site=${site}`);
                continue
            }

            try {
                data = siteFunction.siteObject.getInfo(symbol, attribute);
            }
            catch (ex) {
                Logger.log(`No SITE Object.  Symbol=${symbol}. Attrib=${attribute}. Site=${site}`);
            }

            if (data?.isAttributeSet(attribute)) {
                return data;
            }
        }

        return data;
    }
}

/**
 * @classdesc Used to track lookup times for a stock symbol.
 */
class FinanceSiteLookupStats {
    /**
     * 
     * @param {String} symbol 
     * @param {FinanceWebSite} siteObject
     */
    constructor(symbol, siteObject) {
        this.symbol = symbol;
        this.siteObject = siteObject;
    }

    /**
     * 
     * @param {Number} timeMs 
     * @returns {FinanceSiteLookupStats}
     */
    setSearchTime(timeMs) {
        this.timeMs = timeMs;
        return this;
    }

    /**
     * 
     * @param {StockAttributes} stockAttributes 
     * @returns {FinanceSiteLookupStats}
     */
    setAttributes(stockAttributes) {
        this.price = stockAttributes.stockPrice;
        this.name = stockAttributes.stockName;
        this.yield = stockAttributes.yieldPct;

        return this;
    }
}

/**
 * Returns a diagnostic list of 3rd party stock lookup info.
 * @returns {any[][]}
 */
function cacheFinanceTest() {                               // skipcq:  JS-0128
    const tester = new CacheFinanceTest();

    return tester.execute();
}

/**
 * @classdesc executes 3rd party data lookup tests.
 */
class CacheFinanceTest {
    constructor() {
        this.cacheTestRun = new CacheFinanceTestRun();
    }

    /**
     * Run the tests.
     * @returns {any[][]}
     */
    execute() {
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "NYSEARCA:VOO");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "NEO:CJP");
        this.cacheTestRun.run("YahooApi", YahooApi.getInfo, "NASDAQ:VTC", "PRICE");
        this.cacheTestRun.run("YahooApi", YahooApi.getInfo, "NEO:CJP", "PRICE");

        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NYSEARCA:VOO");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:FTN-A");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NEO:CJP");

        this.cacheTestRun.run("GoogleWebSiteFinance", GoogleWebSiteFinance.getInfo, "NEO:CJP");
        this.cacheTestRun.run("GoogleWebSiteFinance", GoogleWebSiteFinance.getInfo, "NYSEARCA:VOO");

        this.cacheTestRun.run("Finnhub", FinnHub.getInfo, "NYSEARCA:VOO", "PRICE");
        this.cacheTestRun.run("AlphaVantage", AlphaVantage.getInfo, "NYSEARCA:VOO", "PRICE");
        this.cacheTestRun.run("AlphaVantage", AlphaVantage.getInfo, "CURRENCY:USDEUR", "PRICE");
        this.cacheTestRun.run("TwelveData", TwelveData.getInfo, "NYSEARCA:VOO", "PRICE");
        this.cacheTestRun.run("TwelveData", TwelveData.getInfo, "NYSEARCA:VOO", "NAME");
        this.cacheTestRun.run("TwelveData", TwelveData.getInfo, "CURRENCY:USDEUR", "PRICE");

        return this.cacheTestRun.getTestRunResults();
    }
}

/**
 * @classdesc track test results for 3rd party site tests.
 */
class CacheFinanceTestRun {
    constructor() {
        /** @type {CacheFinanceTestStatus[]} */
        this.testRuns = [];
    }

    /**
     * Run one test.
     * @param {String} serviceName 
     * @param {*} func 
     * @param {String} symbol 
     * @param {String} attribute
     * @param {String} type
     */
    run(serviceName, func, symbol, attribute = "ALL", type = "ETF") {
        const result = new CacheFinanceTestStatus(serviceName, symbol);
        try {
            /** @type {StockAttributes} */
            let data = func(symbol, attribute, type);

            if (!(data instanceof StockAttributes)) {
                const myData = new StockAttributes;
                if (attribute === "PRICE") {
                    myData.stockPrice = data;
                    data = myData;
                }
            }

            result.setStatus("ok")
                .setStockAttributes(data)
                .setTypeLookup(type)
                .setAttributeLookup(attribute);

            if (data.stockName === null && data.stockPrice === null && data.yieldPct === null) {
                result.setStatus("Not Found!")
            }
        }
        catch (ex) {
            result.setStatus(`Error: ${ex.toString()}`);
        }
        result.finishTimer();

        this.testRuns.push(result);
    }

    /**
     * Return results for all tests.
     * @returns {any[][]}
     */
    getTestRunResults() {
        const resultTable = [];

        /** @type {any[]} */
        let row = ["Service", "Symbol", "Status", "Price", "Yield", "Name", "Attribute", "Type", "Run Time(ms)"];
        resultTable.push(row);

        for (const testRun of this.testRuns) {
            row = [];

            row.push(testRun.serviceName,
                testRun.symbol,
                testRun.status,
                testRun.stockAttributes.stockPrice,
                testRun.stockAttributes.yieldPct,
                testRun.stockAttributes.stockName,
                testRun._attributeLookup,
                testRun.typeLookup,
                testRun.runTime);

            resultTable.push(row);
        }

        return resultTable;
    }
}

/**
 * @classdesc Individual test results and tracking.
 */
class CacheFinanceTestStatus {
    constructor(serviceName = "", symbol = "") {
        this._serviceName = serviceName;
        this._symbol = symbol;
        this._stockAttributes = new StockAttributes();
        this._startTime = Date.now()
        this._typeLookup = "";
        this._attributeLookup = "";
        this._runTime = 0;
    }

    get serviceName() {
        return this._serviceName;
    }
    get symbol() {
        return this._symbol;
    }
    get value() {
        return this._value;
    }
    get status() {
        return this._status;
    }
    get runTime() {
        return this._runTime;
    }
    get stockAttributes() {
        return this._stockAttributes;
    }
    get typeLookup() {
        return this._typeLookup;
    }
    get attributeLookup() {
        return this._attributeLookup;
    }

    /**
     * 
     * @param {String} val 
     * @returns {CacheFinanceTestStatus}
     */
    setServiceName(val) {
        this._serviceName = val;
        return this;
    }

    /**
     * 
     * @param {String} val 
     * @returns {CacheFinanceTestStatus}
     */
    setSymbol(val) {
        this._symbol = val;
        return this;
    }

    /**
     * 
     * @param {any} val 
     * @returns {CacheFinanceTestStatus}
     */
    setValue(val) {
        this._value = val;
        return this;
    }

    /**
     * 
     * @param {String} val 
     * @returns {CacheFinanceTestStatus}
     */
    setStatus(val) {
        this._status = val;
        return this;
    }

    /**
     * 
     * @param {String} val 
     * @returns {CacheFinanceTestStatus}
     */
    setTypeLookup(val) {
        this._typeLookup = val;
        return this;
    }

    /**
     * 
     * @param {String} val 
     * @returns {CacheFinanceTestStatus}
     */
    setAttributeLookup(val) {
        this._attributeLookup = val;
        return this;
    }

    /**
     * 
     * @param {StockAttributes} val 
     * @returns {CacheFinanceTestStatus}
     */
    setStockAttributes(val) {
        this._stockAttributes = val;
        return this;
    }

    /**
     * 
     * @returns {CacheFinanceTestStatus}
     */
    finishTimer() {
        this._runTime = Date.now() - this._startTime;
        return this;
    }
}



/**
 * @classdesc Concrete implementations for each finance website access.
 */
class FinanceWebSites {
    /**
     * All finance website lookup objects should be defined here.
     * All finance Website Objects must implement the method "getInfo(symbol)" 
     * The getInfo() method must return an instance of "StockAttributes"
     */
    constructor() {
        this.siteList = [
            new FinanceWebSite("YahooApi", YahooApi),
            new FinanceWebSite("GoogleWebSiteFinance", GoogleWebSiteFinance),
            new FinanceWebSite("FinnHub", FinnHub),
            new FinanceWebSite("Globe", GlobeAndMail),
            new FinanceWebSite("Yahoo", YahooFinance),
            new FinanceWebSite("TwelveData", TwelveData),
            new FinanceWebSite("AlphaVantage", AlphaVantage)
        ];

        /** @property {Map<String, FinanceWebSite>} */
        this.siteMap = new Map();
        for (const site of this.siteList) {
            this.siteMap.set(site.siteName, site);
        }
    }

    /**
     * 
     * @returns {FinanceWebSite[]}
     */
    get() {
        return this.siteList;
    }

    /**
     * Get info for running function to get info about a stock using a specfic service name
     * @param {String} name - Name of search to find object containing class info to make function call.
     * @returns {FinanceWebSite}
     */
    getByName(name) {
        const siteInfo = this.siteMap.get(name.toUpperCase());

        return (siteInfo === undefined) ? null : siteInfo;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static getTickerCountryCode(symbol) {
        let exchange = "";
        let countryCode = "";

        if (symbol.indexOf(":") > 0) {
            const parts = symbol.split(":");
            exchange = parts[0].toUpperCase();
        }

        switch (exchange) {
            case "NASDAQ":
            case "NYSEARCA":
            case "NYSE":
            case "NYSEAMERICAN":
            case "OPRA":
            case "OTCMKTS":
            case "BATS":
                countryCode = "us";
                break;
            case "CVE":
            case "TSE":
            case "TSX":
            case "TSXV":
            case "NEO":
                countryCode = "ca";
                break;
            case "SGX":
                countryCode = "sg";
                break;
            case "CURRENCY":
                countryCode = "fx";
                break;
            case "CF":      //  a Mutual fund
                countryCode = "mut";
                break;
            default:
                countryCode = "ca";     //  We the north!
                break;
        }

        return countryCode;
    }

    /**
    * 
    * @param {String} symbol 
    * @returns {String}
    */
    static getBaseTicker(symbol) {
        let modifiedSymbol = symbol;
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const symbolParts = symbol.split(":");

            modifiedSymbol = symbolParts[1];
        }
        return modifiedSymbol;
    }

    /**
     * Get script property using key.  If not found, returns null.
     * @param {String} propertyKey 
     * @returns {any}
     */
    static getApiKey(propertyKey) {
        const scriptProperties = PropertiesService.getScriptProperties();

        const myData = scriptProperties.getProperty(propertyKey);

        return myData;
    }
}

/**
 * @classdesc - defines an instance of an object to perform getInfo(symbol)
 */
class FinanceWebSite {
    /**
     * 
     * @param {String} siteName 
     * @param {object} webSiteJsClass - Points to class for getting finance data.
     */
    constructor(siteName, webSiteJsClass) {
        this.siteName = siteName.toUpperCase();
        this._siteObject = webSiteJsClass;
    }

    set siteName(siteName) {
        this._siteName = siteName;
    }
    get siteName() {
        return this._siteName;
    }

    set siteObject(siteObject) {
        this._siteObject = siteObject;
    }
    get siteObject() {
        return this._siteObject;
    }
}

/**
 * Get/Set data about stocks/ETFs.
 */
class StockAttributes {
    constructor() {
        this._yieldPct = null;
        this._stockName = null;
        this._stockPrice = null;
    }

    get yieldPct() {
        return this._yieldPct;
    }
    set yieldPct(value) {
        if (value !== null) {
            this._yieldPct = Math.round(value * 10000) / 10000;
        }
    }

    get stockPrice() {
        return this._stockPrice;
    }
    set stockPrice(value) {
        if (value !== null) {
            this._stockPrice = Math.round(value * 100) / 100;
        }
    }

    set exchangeRate(value) {
        //  We need 4 decimal places for currency, so setting stockPrice was 
        //  rounding to 2 places - so we have a 'special' price setter for exchange rates.
        if (value !== null) {
            this._stockPrice = Math.round(value * 10000) / 10000;
        }
    }
    get exchangeRate() {
        return this._stockPrice;
    }

    get stockName() {
        return this._stockName;
    }
    set stockName(value) {
        this._stockName = value;
    }

    /**
     * 
     * @param {String} attribute 
     * @returns {any}
     */
    getValue(attribute) {
        switch (attribute) {
            case "PRICE":
                return (this.stockPrice === null) ? '' : this.stockPrice;

            case "YIELDPCT":
                return (this.yieldPct === null) ? '' : this.yieldPct;

            case "NAME":
                return (this.stockName === null) ? '' : this.stockName;

            default:
                return '#N/A';
        }
    }

    /**
     * 
     * @param {String} attribute 
     * @returns {Boolean}
     */
    isAttributeSet(attribute) {
        switch (attribute) {
            case "PRICE":
                {
                    const retVal = this.stockPrice !== null && !Number.isNaN(this.stockPrice) && this.stockPrice !== 0;
                    Logger.log(`price=${this.stockPrice}. Is Valid=${retVal}`);
                    return retVal;
                }

            case "YIELDPCT":
                return this.yieldPct !== null;

            case "NAME":
                return this.stockName !== null;

            default:
                return false;
        }
    }
}


/**
 * @classdesc Lookup for Yahoo site.
 */
class YahooFinance {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        const URL = YahooFinance.getURL(symbol, attribute);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            Logger.log(`FAILED -> getInfo:  ${symbol}.  URL = ${URL}. Err=${ex.toString()}`);
            return new StockAttributes();
        }

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        return YahooFinance.parseResponse(html, symbol);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
        if (FinanceWebSites.getTickerCountryCode(symbol) === "fx") {
            return "";
        }

        return `https://finance.yahoo.com/quote/${YahooFinance.getTicker(symbol)}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return "";
    }

    /**
     * 
     * @param {String} html 
     * @param {String} symbol
     * @returns {StockAttributes}
     */
    static parseResponse(html, symbol) {
        const data = new StockAttributes();

        if (symbol === '') {
            return data;
        }

        //  skipcq:  JS-0097
        let dividendPercent = html.match('title="Yield">Yield</span> <[^>]*>(\\d{0,5}\.?\\d{0,4})');
        if (dividendPercent !== null && dividendPercent.length === 2) {
            const tempPct = dividendPercent[1];
            Logger.log(`Yahoo. Stock=${symbol}. PERCENT=${tempPct}`);

            data.yieldPct = Number.parseFloat(tempPct) / 100;

            if (Number.isNaN(data.yieldPct)) {
                data.yieldPct = null;
            }
        }

        //  skipcq:  JS-0097
        const priceMatch = html.match('qsp-price">(\\d{0,5}\.?\\d{0,4})');
        if (priceMatch !== null && priceMatch.length === 2) {
            const tempPrice = priceMatch[1];
            Logger.log(`Yahoo. Stock=${symbol}.PRICE=${tempPrice}`);

            data.stockPrice = Number.parseFloat(tempPrice);

            if (Number.isNaN(data.stockPrice)) {
                data.stockPrice = null;
            }
        }

        const baseSymbol = YahooFinance.getTicker(symbol);
        // skipcq: JS-0097
        const nameRegex = new RegExp(`<title>(.+?)\(${baseSymbol}\)`);
        const nameMatch = html.match(nameRegex);
        if (nameMatch !== null && nameMatch.length > 1) {
            data.stockName = nameMatch[1].endsWith("(") ? nameMatch[1].slice(0, -1) : nameMatch[1];
            Logger.log(`Yahoo. Stock=${symbol}.NAME=${data.stockName}`);
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static getTicker(symbol) {
        let modifiedSymbol = symbol;
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const symbolParts = symbol.split(":");

            modifiedSymbol = symbolParts[1];
            if (symbolParts[0] === "TSE")
                modifiedSymbol = `${symbolParts[1]}.TO`;
            if (symbolParts[0] === "SGX")
                modifiedSymbol = `${symbolParts[1]}.SI`;
            if (symbolParts[0] === "NEO")
                modifiedSymbol = `${symbolParts[1]}.NE`;
            if (symbolParts[0] === "AS")
                modifiedSymbol = `${symbolParts[1]}.AS`;
            if (symbolParts[0] === "MI")
                modifiedSymbol = `${symbolParts[1]}.MI`;

        }
        return modifiedSymbol;
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        return null;
    }
}

class YahooApi {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        const URL = YahooApi.getURL(symbol, attribute);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        return YahooApi.parseResponse(html, symbol);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {String}
     */
    static getURL(symbol, attribute) {
        if (FinanceWebSites.getTickerCountryCode(symbol) === "fx") {
            return "";
        }

        if (attribute !== "PRICE" && attribute !== "NAME") {
            return "";
        }

        return `https://query1.finance.yahoo.com/v8/finance/chart/${YahooApi.getTicker(symbol)}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return "";
    }

    /**
     * 
     * @param {String} html 
     * @param {String} symbol
     * @returns {StockAttributes}
     */
    static parseResponse(html, symbol) {
        const stockData = new StockAttributes();

        if (symbol === '') {
            return stockData;
        }

        try {
            const data = JSON.parse(html);
            if (data?.chart?.result.length > 0) {
                const regularMarketPrice = data.chart.result[0].meta.regularMarketPrice;

                stockData.stockPrice = parseFloat(regularMarketPrice);

                if (Number.isNaN(stockData.stockPrice)) {
                    stockData.stockPrice = null;
                }

                stockData.stockName = data.chart.result[0].meta.longName;
            }
        }
        catch (ex) {
            Logger.log(`Failed to parse JSON: ${symbol}`);
        }

        return stockData;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static getTicker(symbol) {
        return YahooFinance.getTicker(symbol);
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        return null;
    }
}

/**
 * @classdesc Lookup for Globe and Mail website.
 */
class GlobeAndMail {
    /**
     * Only gets dividend yield.
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        const URL = GlobeAndMail.getURL(symbol, attribute);

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }

        return GlobeAndMail.parseResponse(html, symbol);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
        if (FinanceWebSites.getTickerCountryCode(symbol) === "fx") {
            return "";
        }

        //  Mutual fund
        if (FinanceWebSites.getTickerCountryCode(symbol) === "mut") {
            return `https://www.theglobeandmail.com/investing/markets/funds/${GlobeAndMail.getTicker(symbol)}`;
        }

        return `https://www.theglobeandmail.com/investing/markets/stocks/${GlobeAndMail.getTicker(symbol)}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return "";
    }

    /**
     * 
     * @param {String} html 
     * @returns {StockAttributes}
     */
    static parseResponse(html) {
        const data = new StockAttributes();

        //  Get the dividend yield.
        let parts = html.match(/.name="dividendYieldTrailing".*?value="(\d{0,4}\.?\d{0,4})%/);

        if (parts === null) {
            parts = html.match(/.name=\\"dividendYieldTrailing\\".*?value=\\"(\d{0,4}\.?\d{0,4})%/);
        }

        if (parts === null) {
            parts = html.match(/"distributionYield" type="percent" value="(\d{0,4}\.?\d{0,4})%"/);
        }

        if (parts !== null && parts.length === 2) {
            const tempPct = parts[1];

            const parsedValue = Number.parseFloat(tempPct) / 100;

            if (!Number.isNaN(parsedValue)) {
                data.yieldPct = parsedValue;
            }
        }

        //  Get the name.
        parts = html.match(/."symbolName":"(.*?)"/);
        if (parts !== null && parts.length === 2) {
            data.stockName = parts[1];
        }


        //  Get the price.
        parts = html.match(/"lastPrice" value="(\d{0,4}\.?\d{0,4})"/);

        if (parts !== null && parts.length === 2) {

            const parsedValue = Number.parseFloat(parts[1]);
            if (!Number.isNaN(parsedValue)) {
                data.stockPrice = parsedValue;
            }
        }

        return data;
    }

    /**
     * Clean up ticker symbol for use in Globe and Mail lookups.
     * @param {String} symbol 
     * @returns {String}
     */
    static getTicker(symbol) {
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const parts = symbol.split(":");

            switch (parts[0].toUpperCase()) {
                case "TSE":
                    symbol = parts[1];
                    if (parts[1].includes(".")) {
                        symbol = parts[1].replace(".", "-");
                    }
                    else if (parts[1].includes("-")) {
                        const prefShare = parts[1].split("-");
                        symbol = `${prefShare[0]}-PR-${prefShare[1]}`;
                    }
                    symbol = `${symbol}-T`;
                    break;
                case "NEO":
                    symbol = `${parts[1]}-NE`;
                    break;
                case "NYSEARCA":
                case "BATS":
                    symbol = `${parts[1]}-A`;
                    break;
                case "NASDAQ":
                    symbol = `${parts[1]}-Q`;
                    break;
                case "CF":
                    symbol = `${parts[1]}.CF`;
                    break;
                default:
                    symbol = '#N/A';
            }
        }

        return symbol;
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        return null;
    }
}



/**
 * @classdesc Uses FINNHUB Rest API.  Requires a script setting for the API key.
 * Set key name as FINNHUB_API_KEY
 */
class FinnHub {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns 
     */
    static getInfo(symbol, attribute = "PRICE") {
        let data = new StockAttributes();

        if (attribute !== "PRICE") {
            Logger.log(`Finnhub.  Only PRICE is supported: ${symbol}, ${attribute}`);
            return data;
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (countryCode !== "us") {
            Logger.log(`FinnHub --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const URL = FinnHub.getURL(symbol, attribute, FinnHub.getApiKey());
        Logger.log(`getInfo:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();
            data = FinnHub.parseResponse(jsonStr, symbol, attribute);
        }
        catch (ex) {
            return data;
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @param {String} API_KEY
     * @returns {String}
     */
    static getURL(symbol, attribute, API_KEY = null) {
        if (attribute !== "PRICE") {
            return "";
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (countryCode !== "us") {
            return "";
        }

        if (API_KEY === null) {
            return "";
        }

        return `https://finnhub.io/api/v1/quote?symbol=${FinanceWebSites.getBaseTicker(symbol)}&token=${API_KEY}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return FinanceWebSites.getApiKey("FINNHUB_API_KEY");
    }

    /**
     * 
     * @param {String} jsonStr 
     * @param {String} _symbol
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static parseResponse(jsonStr, _symbol, attribute) {
        const data = new StockAttributes();

        const hubData = JSON.parse(jsonStr);
        if (attribute === "PRICE")
            data.stockPrice = hubData.c;

        return data;
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        //  Basic throttle check
        const limits = [
            new ThresholdPeriod("SECOND", 30),
            new ThresholdPeriod("MINUTE", 60)
        ];

        return new SiteThrottle("FINNHUB", limits);
    }
}

class AlphaVantage {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute = "PRICE") {
        let data = new StockAttributes();

        if (attribute !== "PRICE") {
            Logger.log(`AlphaVantage.  Only PRICE is supported: ${symbol}, ${attribute}`);
            return data;
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (!(countryCode === "us" || countryCode === "fx")) {
            Logger.log(`AlphaVantage --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const apiKey = AlphaVantage.getApiKey();
        const URL = AlphaVantage.getURL(symbol, attribute, apiKey);
        Logger.log(`getInfo: AlphaVantage  ${symbol}.  URL = ${URL}.  Key = ${apiKey}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();
            data = AlphaVantage.parseResponse(jsonStr, symbol, attribute);
        }
        catch (ex) {
            return data;
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @param {String} API_KEY
     * @returns {String}
     */
    static getURL(symbol, attribute, API_KEY = null) {
        if (API_KEY === null) {
            return "";
        }

        if (attribute !== "PRICE") {
            return "";
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (!(countryCode === "us" || countryCode === "fx")) {
            return "";
        }
        const symbolParts = symbol.split(":");

        if (countryCode === "fx") {
            const fromCurrency = symbolParts[1].substring(0, 3);
            const toCurrency = symbolParts[1].substring(3, 6);

            return `https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=${fromCurrency}&to_currency=${toCurrency}&apikey=${API_KEY}`;
        }
        else {
            return `https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=${FinanceWebSites.getBaseTicker(symbol)}&apikey=${API_KEY}`;
        }
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return FinanceWebSites.getApiKey("ALPHA_VANTAGE_API_KEY");
    }

    /**
     * 
     * @param {String} jsonStr 
     * @param {String} symbol
     * @param {String} _attribute
     * @returns {StockAttributes}
     */
    static parseResponse(jsonStr, symbol, _attribute) {
        Logger.log(`content=${jsonStr}`);

        const data = new StockAttributes();
        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        const alphaVantageData = JSON.parse(jsonStr);

        try {
            if (countryCode === "fx") {
                data.exchangeRate = alphaVantageData["Realtime Currency Exchange Rate"]["5. Exchange Rate"];
            }
            else {
                data.stockPrice = alphaVantageData["Global Quote"]["05. price"];
            }

            Logger.log(`Price=${data.stockPrice}`);
        }
        catch (ex) {
            Logger.log(`AlphaVantage JSON Parse Error (looking for ${countryCode}. err=${ex}).`);
        }

        return data;
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        //  Basic throttle check
        const limits = [
            new ThresholdPeriod("DAY", 25)
        ];

        return new SiteThrottle("ALPHAVANTAGE", limits);
    }
}

/**
 * @classdesc Lookup for GOOGLE Finance site.
 * weirdly, GOOGLEFINANCE() will fail for some stocks, but work on the website.
 */
class GoogleWebSiteFinance {
    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        const URL = GoogleWebSiteFinance.getURL(symbol, attribute);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        return GoogleWebSiteFinance.parseResponse(html, symbol);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {String}
     */
    static getURL(symbol, attribute) {
        if (attribute === "YIELDPCT" && FinanceWebSites.getTickerCountryCode(symbol) === "us") {
            //  This site is very bad at yields for u.s.
            return "";
        }

        return `https://www.google.com/finance/quote/${GoogleWebSiteFinance.getTicker(symbol)}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return "";
    }

    /**
     * 
     * @param {String} html 
     * @param {String} symbol
     * @returns {StockAttributes}
     */
    static parseResponse(html, symbol) {
        const data = new StockAttributes();

        if (symbol === '') {
            return data;
        }

        if (html.indexOf("We couldn't find any match for your search.") !== -1) {
            Logger.log(`www.google.com/finance:  We couldn't find any match for your search. symbol=${symbol}`);
            return data;
        }

        data.yieldPct = GoogleWebSiteFinance.extractYieldPct(html, symbol);
        data.stockName = GoogleWebSiteFinance.extractStockName(html, symbol);
        if (FinanceWebSites.getTickerCountryCode(symbol) === "fx") {
            data.exchangeRate = GoogleWebSiteFinance.extractStockPrice(html, symbol);
        }
        else {
            data.stockPrice = GoogleWebSiteFinance.extractStockPrice(html, symbol);
        }

        Logger.log(`Google. Stock=${symbol}. PERCENT=${data.yieldPct}. NAME=${data.stockName}. PRICE=${data.stockPrice}`);

        return data;
    }

    /**
     * 
     * @param {String} html 
     * @param {String} _symbol 
     * @returns {Number}
     */
    static extractYieldPct(html, _symbol) {
        let data = null;
        //  skipcq: JS-0097
        const dividendPercent = html.match(/Dividend yield.+?(\d{0,4}\.?\d{0,4})%<\/div>/);

        if (dividendPercent !== null && dividendPercent.length > 1) {
            const tempPct = dividendPercent[1];

            data = Number.parseFloat(tempPct) / 100;

            if (Number.isNaN(data)) {
                data = null;
            }
        }
        return data;
    }

    /**
     * 
     * @param {String} html 
     * @param {String} _symbol 
     * @returns {Number}
     */
    static extractStockPrice(html, _symbol) {
        let data = null;
        //  skipcq: JS-0097
        const priceMatch = html.match(/data-last-price="(\d{0,7}\.*\d{0,20})"/);
        if (priceMatch !== null && priceMatch.length > 1) {
            const tempPrice = priceMatch[1];

            data = Number.parseFloat(tempPrice);

            if (Number.isNaN(data)) {
                data = null;
            }
        }

        return data;
    }

    /**
     * 
     * @param {String} html 
     * @param {String} symbol 
     * @returns {String}
     */
    static extractStockName(html, symbol) {
        let data = null;
        const baseSymbol = GoogleWebSiteFinance.getTicker(symbol);
        const stockNameParts = baseSymbol.split(":");

        if (stockNameParts.length > 1) {
            const stock = stockNameParts[0];

            //  skipcq: JS-0097
            const nameRegex = new RegExp(`<title>(.+?)\(${stock}\)`);
            const nameMatch = html.match(nameRegex);
            if (nameMatch !== null && nameMatch.length > 1) {
                data = nameMatch[1].endsWith("(") ? nameMatch[1].slice(0, -1) : nameMatch[1];
            }
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static getTicker(symbol) {
        let modifiedSymbol = symbol;
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const symbolParts = symbol.split(":");

            if (FinanceWebSites.getTickerCountryCode(symbol) === "fx") {
                const fromCurrency = symbolParts[1].substring(0, 3);
                const toCurrency = symbolParts[1].substring(3, 6);
                modifiedSymbol = `${fromCurrency}-${toCurrency}?hl=en`;
            }
            else {
                modifiedSymbol = `${symbolParts[1]}:${symbolParts[0]}`;
            }
        }
        return modifiedSymbol;
    }


    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        return null;
    }
}

class TwelveData {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute = "PRICE") {
        let data = new StockAttributes();

        if (attribute !== "PRICE" && attribute !== "NAME") {
            Logger.log(`TwelveData.  Only PRICE/NAME is supported: ${symbol}, ${attribute}`);
            return data;
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (!(countryCode === "us" || countryCode === "fx")) {
            Logger.log(`TwelveData --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const apiKey = TwelveData.getApiKey();
        const URL = TwelveData.getURL(symbol, attribute, apiKey);
        Logger.log(`getInfo: TwelveData  ${symbol}.  URL = ${URL}.  Key = ${apiKey}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();
            data = TwelveData.parseResponse(jsonStr, symbol, attribute);
        }
        catch (ex) {
            return data;
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @param {String} API_KEY
     * @returns {String}
     */
    static getURL(symbol, attribute, API_KEY = null) {
        if (API_KEY === null) {
            return "";
        }

        if (attribute !== "PRICE" && attribute !== "NAME") {
            return "";
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (!(countryCode === "us" || countryCode === "fx")) {
            return "";
        }

        let twelveDataSymbol = "";
        if (countryCode === "fx") {
            const symbolParts = symbol.split(":");
            const fromCurrency = symbolParts[1].substring(0, 3);
            const toCurrency = symbolParts[1].substring(3, 6);

            twelveDataSymbol = `${fromCurrency}/${toCurrency}`;
        }
        else {
            twelveDataSymbol = FinanceWebSites.getBaseTicker(symbol);
        }

        return `https://api.twelvedata.com/quote?symbol=${twelveDataSymbol}&apikey=${API_KEY}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return FinanceWebSites.getApiKey("TWELVE_DATA_API_KEY");
    }

    /**
     * 
     * @param {String} jsonStr 
     * @param {String} symbol
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static parseResponse(jsonStr, symbol, attribute) {
        Logger.log(`content=${jsonStr}`);

        const data = new StockAttributes();
        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        const twelveData = JSON.parse(jsonStr);

        try {
            if (attribute === "NAME") {
                data.stockName = twelveData.name;
                Logger.log(`TwelveData. Name=${data.stockName}`);
            }
            else if (attribute === "PRICE") {
                if (countryCode === "fx") {
                    data.exchangeRate = twelveData.close;
                }
                else {
                    data.stockPrice = twelveData.close;
                }
                Logger.log(`TwelveData. Price=${data.stockPrice}`);
            }
        }
        catch (ex) {
            Logger.log(`TwelveData JSON Parse Error (looking for ${countryCode}. err=${ex}).`);
        }

        return data;
    }

    /**
     * Get an instance of the throttling object to query if the web limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        //  Basic throttle check
        const limits = [
            new ThresholdPeriod("MINUTE", 8),
            new ThresholdPeriod("DAY", 800)
        ];

        return new SiteThrottle("TWELVEDATA", limits);
    }
}


/**
 * @classdesc Multi-purpose functions used within the cache finance custom functions.
 */
class CacheFinanceUtils {                       // skipcq: JS-0128
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
     * @param {Number} cacheSeconds
     */
    static bulkShortCachePut(symbols, attribute, newCacheData, cacheSeconds) {
        if (symbols.length === 0 || newCacheData.length === 0) {
            return;
        }

        const cacheKeyList = CacheFinanceUtils.createCacheKeyList(symbols, attribute);
        CacheFinanceUtils.putFinanceValuesIntoShortCache(cacheKeyList, newCacheData, cacheSeconds);
    }

    /**
     * Create unique key/value key, filter out bad data, save to script settings (long cache).
     * @param {any[]} symbols 
     * @param {String} attribute 
     * @param {any[]} cacheData 
     */
    static bulkLongCachePut(symbols, attribute, cacheData, daysToHold = 7) {
        const cacheKeys = CacheFinanceUtils.createCacheKeyList(symbols, attribute);
        const newCacheKeys = [];
        const newCacheData = [];

        cacheKeys.forEach((key, i) => {
            if (CacheFinanceUtils.isValidGoogleValue(cacheData[i])) {
                newCacheKeys.push(key);
                newCacheData.push(cacheData[i])
            }
        });

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

        cacheKeys.forEach(key => {
            const parsedData = data[key] === undefined ? null : JSON.parse(data[key]);
            cachedDataList.push(parsedData);
        });

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
        let updateCounter = 0;

        for (let i = 0; i < cacheKeys.length; i++) {
            if (CacheFinanceUtils.isValidGoogleValue(newCacheData[i])) {
                bulkData[cacheKeys[i]] = JSON.stringify(newCacheData[i]);
                updateCounter++;
            }
        }

        if (updateCounter > 0) {
            const shortCache = CacheService.getScriptCache();
            shortCache.putAll(bulkData, cacheSeconds);
        }
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {String}
     */
    static makeCacheKey(symbol, attribute) {
        return `${attribute.toUpperCase()}|${symbol.toUpperCase()}`;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {String}
     */
    static makeIgnoreSiteCacheKey(symbol, attribute) {
        return `IGNORE|${CacheFinanceUtils.makeCacheKey(symbol, attribute)}`;
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
        return value !== null && value !== undefined && value !== "#N/A" && value !== '#ERROR!' && value !== '';
    }

    /**
     * When you request a single column of data from getRange(), it is still a double array.
     * Convert to single array for reguar array processing.
     * @param {any[][]} doubleArray 
     * @param {Number} columnNumber 
     * @returns {any[]}
     */
    static convertRowsToSingleArray(doubleArray, columnNumber = 0) {
        if (!Array.isArray(doubleArray)) {
            return doubleArray;
        }

        return doubleArray.map(item => item[columnNumber]);
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

/**
 * @classdesc Used for tracking throttle use limits for a stock site.
 */
class SiteThrottle {            // skipcq:  JS-0128
    /**
     * @param {String} siteID
     * @param {ThresholdPeriod[]} thresholds 
     */
    constructor(siteID, thresholds) {
        this.siteID = siteID;
        this.thresholds = thresholds;
        this.periodKeys = [];
        this.periodCount = [];
    }

    /**
     * @returns {Boolean} - true is ok to make request
     */
    checkAndIncrement() {
        let inLimit = true;

        //  If this is the first time called, we must do the time intensive read from cache
        //  to get the current api call count for each period we have limits for.
        if (this.periodKeys.length === 0) {
            [this.periodKeys, this.periodCount] = SiteThrottle.getCurrentThresholds(this.thresholds, this.siteID);
        }

        //  Will we exceed any thresholds?
        for (let i = 0; i < this.thresholds.length; i++) {
            inLimit = this.periodCount[i] + 1 < this.thresholds[i].maxPerPeriod;

            if (!inLimit) {
                Logger.log(`Throttle Limit EXCEEDED. ${this.siteID}`);
                break;
            }
        }

        //  If we don't exceed the throttle limits, increment all threshold counters.
        if (inLimit) {
            for (let i = 0; i < this.periodCount.length; i++) {
                this.periodCount[i]++;
            }
        }

        return inLimit;
    }

    /**
     * @param {ThresholdPeriod[]} thresholds 
     * @param {String} siteID 
     * @returns {[String[], Number[]]}
     */
    static getCurrentThresholds(thresholds, siteID) {
        let key = "";
        let current = 0;
        const keys = [];
        const limits = [];

        for (const period of thresholds) {
            switch (period.periodName) {
                case "SECOND":
                    key = SiteThrottle.createSecondKey(siteID);
                    current = SiteThrottle.currentForSecond(key);
                    Logger.log(`SECOND Check.  key=${key}. Current=${current.toString()}.`);
                    break;

                case "MINUTE":
                    key = SiteThrottle.createMinuteKey(siteID);
                    current = SiteThrottle.currentForMinute(key);
                    Logger.log(`MINUTE Check.  key=${key}. Current=${current.toString()}.`);
                    break;

                case "DAY":
                    key = SiteThrottle.createDayKey(siteID);
                    current = SiteThrottle.currentForDay(key);
                    Logger.log(`DAY Check.  key=${key}. Current=${current.toString()}.`);

                    break;

                default:
                    throw new Error(`Invalid threshold period ${period.periodName}`);
            }

            //  We save the KEY because it may change before we increment.
            //  We save the current value because re-reading is very time consuming.
            keys.push(key);
            limits.push(current);
        }

        return [keys, limits];
    }

    //  At the end of a batch URL fetch, the THROTTLE stats need to be 
    //  saved to cache.
    update() {
        if (this.periodKeys.length === 0) {
            //  Was never used for this site.
            return;
        }

        for (let i = 0; i < this.thresholds.length; i++) {
            const period = this.thresholds[i];
            switch (period.periodName) {
                case "SECOND":
                    SiteThrottle.updateForSecond(this.periodKeys[i], this.periodCount[i]);
                    break;

                case "MINUTE":
                    SiteThrottle.updateForMinute(this.periodKeys[i], this.periodCount[i]);
                    break;

                case "DAY":
                    SiteThrottle.updateForDay(this.periodKeys[i], this.periodCount[i]);
                    break;

                default:
                    throw new Error(`Invalid threshold period ${period.periodName}`);
            }
        }

        this.periodKeys = [];
        this.periodCount = [];
    }

    /**
     * 
     * @param {String} key 
     * @returns {Number}
     */
    static currentForSecond(key) {
        const shortCache = CacheService.getScriptCache();
        const data = shortCache.get(key);

        return data === null ? 0 : JSON.parse(data);            
    }

    /**
     * Current number of requests in THIS minute.  
     * @param {String} key
     * @returns {Number}
     */
    static currentForMinute(key) {
        const shortCache = CacheService.getScriptCache();
        const data = shortCache.get(key);

        return data === null ? 0 : JSON.parse(data);
    }

    /**
     * 
     * @param {String} key 
     * @param {Number} current 
     */
    static updateForSecond(key, current) {
        const shortCache = CacheService.getScriptCache();
        shortCache.put(key, JSON.stringify(current), 10);        
    }

    /**
     * Add to minute counter.
     * @param {String} key 
     * @param {Number} current 
     */
    static updateForMinute(key, current) {
        const shortCache = CacheService.getScriptCache();
        shortCache.put(key, JSON.stringify(current), 180);
    }

    /**
     * @param {String} key 
     * @param {Number} current 
     */
    static updateForDay(key, current) {
        const longCache = new ScriptSettings();
        longCache.put(key, current, 2);
    }

    /**
     * Current requests made for the day.
     * @param {String} key 
     * @returns {Number}
     */
    static currentForDay(key) {
        const longCache = new ScriptSettings();
        const data = longCache.get(key);

        return data === null ? 0 : data;
    }

    /**
     * 
     * @param {String} siteID 
     * @param {String} intervalName 
     * @param {any} periodNumber - (0-59 -> MINUTE), (0-6 -> DAY) 
     * @returns 
     */
    static makeKey(siteID, intervalName, periodNumber) {
        return `${siteID}:${intervalName}:${periodNumber.toString()}`;
    }

    /**
     * 
     * @param {String} siteID 
     * @returns {String}
     */
    static createSecondKey(siteID) {
        const today = new Date();
        const second = today.getSeconds();

        return SiteThrottle.makeKey(siteID, "SEC", second);        
    }

    /**
     * 
     * @param {String} siteID 
     * @returns {String}
     */
    static createMinuteKey(siteID) {
        const today = new Date();
        const minute = today.getMinutes();

        return SiteThrottle.makeKey(siteID, "MIN", minute);
    }

    /**
     * 
     * @param {String} siteID 
     * @returns {String}
     */
    static createDayKey(siteID) {
        const today = new Date();
        const dayNum = today.getDay();      // Day of the week. 0-6
        return SiteThrottle.makeKey(siteID, "DAY", dayNum);
    }
}


/**
 * @classdesc Used to define a throttle limit.
 */
class ThresholdPeriod {         // skipcq:  JS-0128
    constructor(periodName, maxPerPeriod) {
        this._periodName = periodName;
        this._maxPerPeriod = maxPerPeriod;
    }

    get periodName() {
        return this._periodName;
    }
    set periodName(val) {
        this._periodName = val;
    }
    get maxPerPeriod() {
        return this._maxPerPeriod;
    }
    set maxPerPeriod(val) {
        this._maxPerPeriod = val;
    }
}

