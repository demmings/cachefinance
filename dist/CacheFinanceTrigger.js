
const GOOGLEFINANCE_PARAM_NOT_USED = "##NotSet##";

//  Function only used for testing in google sheets app script.
// skipcq: JS-0128
function testYieldPct() {
    const val = CACHEFINANCE("TSE:FTN-A", "yieldpct");        // skipcq: JS-0128
    Logger.log(`Test CacheFinance FTN-A(yieldpct)=${val}`);
}

function testCacheFinances() {                                  // skipcq: JS-0128
    // const symbols = [["ABC"], ["DEF"], ["GHI"], ["JKL"], ["TSE:FLJA"]];
    // const data = [[11.1], [22.2], [33.3], [44.4], ["#N/A"]];

    const symbols = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("A30:A165").getValues();
    const data = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("E30:E165").getValues();

    const cacheData = CACHEFINANCES(symbols, "PRICE", data);

    const singleSymbols = CacheFinanceUtils.convertRowsToSingleArray(symbols);

    Logger.log(`BULK CACHE TEST Success${cacheData} . ${singleSymbols}`);
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

/**
 * Run test manually within Sheets.
 */
function testCacheTriggerJob() {                // skipcq: JS-0128
    const jobParms = ["Stocks!A8:A140",
        "PRICE",
        "Stocks!C8:C140",
        "Stocks!CB8:CB140",
        1,
        "*",
        "MON-SUN",
        ""];

    const jobSettings = new CacheJob(jobParms);
    CacheTrigger.getJobData(jobSettings);
}

/**
 * Run test manually within Sheets.
 */
function testBulkCache() {                      // skipcq: JS-0128
    const symbols = [["ABC"], ["DEF"], ["GHI"], ["JKL"], ["MNO"]];
    const data = [11.1, 22.2, 33.3, 44.4, 55.5];

    CacheFinanceUtils.bulkShortCachePut(symbols, "PRICE", data);
    const cacheData = CacheFinanceUtils.bulkShortCacheGet(symbols, "PRICE");

    if (JSON.stringify(data) !== JSON.stringify(cacheData)) {
        Logger.log("BULK Cache TEST Fail.");
    }

    Logger.log("BULK CACHE TEST Success");
}

//  Named range in sheet with CacheFinance configurations.
const CACHE_LEGEND = "CACHEFINANCE";

/**
 * Add this to App Script Trigger.
 * It requires a named range in your sheet called 'CACHEFINANCE'
 * @param {GoogleAppsScript.Events.TimeDriven} e 
 */
function CacheFinanceTrigger(e) {                           //  skipcq: JS-0128
    Logger.log("Starting CacheFinanceTrigger()");

    const cacheSettings = new CacheJobSettings();

    //  The trigger ID for THIS job is already disabled.  Send a signal to other
    //  running jobs that THIS trigger is still going.
    CacheJobSettings.signalTriggerRunState(e, true);

    //  Is this job specified in legend.
    /** @type {CacheJob} */
    const jobInfo = cacheSettings.getMyJob(e);

    if (jobInfo === null && (typeof e !== 'undefined')) {
        //  This is a boot job that won't run again.
        cacheSettings.firstRun(e.triggerUid);
        return;
    }

    //  Run job to update finance data.
    CacheTrigger.runJob(jobInfo);

    //  Create new job and ensure all existing jobs are valid.
    cacheSettings.afterRun(jobInfo);

    //  Signal outside trigger that this ID is not running.
    CacheJobSettings.signalTriggerRunState(e, false);
}

/**
 * Add a custom function to your sheet to get the Trigger(s) installed.
 * The customfunction cannot modify the sheet settings for our jobs, so
 * the initial trigger for CacheFinanceTrigger creates the jobs and ends.
 * The actual jobs will run after that.
 * @returns {String[][]}
 * @customfunction
 */
function CacheFinanceBoot() {                       //  skipcq: JS-0128
    if (CacheJobSettings.bootstrapTrigger())
        return [["Trigger Created!"]];
    else
        return [["Trigger Exists"]];
}

/**
 * @classdesc Manage TRIGGERS that run and update stock/etf data from the web.
 */
class CacheJobSettings {
    constructor() {
        this.load(null);
    }

    /**
     * 
     * @param {CacheJob} currentJob 
     */
    load(currentJob) {
        const sheetNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(CACHE_LEGEND);

        if (sheetNamedRange === null) {
            Logger.log("Named Range CACHEFINANCE not found.");
            Logger.log("Each definition line must define:  'SymbolRange', 'Attribute', 'OutputRange', 'GoogleFinanceRange', 'Refresh Minutes', 'Trigger ID'");
            throw new Error("Named Range CACHEFINANCE not defined.");
        }

        /** @type {CacheJob[]} */
        this.jobs = [];

        this.cacheInfo = sheetNamedRange.getValues();

        for (const job of this.cacheInfo) {
            const cacheJob = new CacheJob(job);

            if (currentJob !== null && cacheJob.triggerID === currentJob.triggerID) {
                this.jobs.push(currentJob);
            }
            else {
                this.jobs.push(cacheJob);
            }
        }
    }

    /**
     * 
     * Adds 'CacheFinanceTrigger' function to triggers.
     * @returns {Boolean} - true = created
     */
    static bootstrapTrigger() {
        let missingTrigger = true;
        const validTriggerList = ScriptApp.getProjectTriggers();
        for (const trigger of validTriggerList) {
            // @ts-ignore
            if (trigger.getHandlerFunction().toUpperCase() === 'CACHEFINANCETRIGGER' && !trigger.isDisabled())
                missingTrigger = false;
        }

        if (missingTrigger) {
            Logger.log("Creating BOOTSTRAP Trigger for CacheFinanceTrigger function.")
            ScriptApp
                .newTrigger('CacheFinanceTrigger')
                .timeBased()
                .after(15000)
                .create();

            return true;
        }

        return false;
    }

    /**
     * 
     * @param {String} triggerUid 
     * @returns {void}
     */
    firstRun(triggerUid) {
        //  On first run, it should create a job for every range specified in CACHEFINANCE legend.
        this.validateTriggerIDs();
        this.createMissingTriggers(true);

        //  This is the BOOT trigger ID.
        CacheJobSettings.deleteOldTrigger(triggerUid);
    }

    /**
     * 
     * @param {CacheJob} jobInfo 
     * @returns {void}
     */
    afterRun(jobInfo) {
        //  Reload job table in case another trigger recently updated.
        this.load(jobInfo);

        CacheJobSettings.deleteOldTrigger(jobInfo.triggerID);       //  Delete myself
        jobInfo.triggerID = "";
        CacheJobSettings.cleanupDisabledTriggers();                 //  Delete triggers that ran, but not cleaned up.    
        this.validateTriggerIDs();
        this.createMissingTriggers(false);
    }

    /**
     * Find trigger that are disabled and not associated with anything running.
     * @returns {void}
     */
    static cleanupDisabledTriggers() {
        const validTriggerList = ScriptApp.getProjectTriggers();
        for (const trigger of validTriggerList) {
            // @ts-ignore
            if (trigger.getHandlerFunction().toUpperCase() === 'CACHEFINANCETRIGGER' && trigger.isDisabled() &&
                !CacheJobSettings.isDisabledTriggerStillRunning(trigger.getUniqueId())) {
                Logger.log("Trigger CLEANUP.  Deleting disabled trigger.");
                ScriptApp.deleteTrigger(trigger);
            }
        }
    }

    /**
     * 
     * @param {GoogleAppsScript.Events.TimeDriven} e 
     * @param {Boolean} alive 
     * @returns {void}
     */
    static signalTriggerRunState(e, alive) {
        if (typeof e === 'undefined') {
            Logger.log("Trigger ID unknown.");
            return;
        }

        const key = CacheFinanceUtils.makeCacheKey("ALIVE", e.triggerUid);
        const shortCache = CacheService.getScriptCache();
        if (alive) {
            shortCache.put(key, "ALIVE");
        }
        else if (shortCache.get(key) !== null) {
            shortCache.remove(key);
        }
    }

    /**
     * Mark the job triggerID if not a valid trigger (in our job table.)
     * @returns {void}
     */
    validateTriggerIDs() {
        Logger.log("Starting validateTriggerIDs()");

        const validTriggerList = ScriptApp.getProjectTriggers();

        for (const job of this.jobs) {
            if (job.triggerID === "")
                continue;

            let good = false;
            for (const validID of validTriggerList) {
                if (job.triggerID === validID.getUniqueId()) {
                    good = true;
                    break;
                }
            }

            if (!good) {
                //  It could be running.
                if (CacheJobSettings.isDisabledTriggerStillRunning(job.triggerID)) {
                    continue;
                }

                Logger.log(`Invalid Trigger ID=${job.triggerID}`);
                job.triggerID = "";
            }
            else {
                Logger.log(`Valid Trigger ID=${job.triggerID}`);
            }
        }
    }

    /**
     * 
     * @param {String} triggerID 
     * @returns {Boolean}
     */
    static isDisabledTriggerStillRunning(triggerID) {
        const shortCache = CacheService.getScriptCache();

        const key = CacheFinanceUtils.makeCacheKey("ALIVE", triggerID);
        if (shortCache.get(key) !== null) {
            Logger.log(`Trigger ID=${triggerID} is RUNNING!`);
            return true;
        }

        return false;
    }

    /**
     * 
     * @param {any} triggerID 
     * @returns {void}
     */
    static deleteOldTrigger(triggerID) {
        const triggers = ScriptApp.getProjectTriggers();

        let triggerObject = null;
        for (const item of triggers) {
            if (item.getUniqueId() === triggerID)
                triggerObject = item;
        }

        if (triggerObject !== null) {
            Logger.log(`DELETING Trigger: ${triggerID}`);
            ScriptApp.deleteTrigger(triggerObject);
        }
        else {
            Logger.log(`Failed to locate trigger to delete: ${triggerID}`);
        }
    }

    /**
     * 
     * @param {Boolean} runAsap 
     * @returns {void}
     */
    createMissingTriggers(runAsap) {
        for (const job of this.jobs) {
            if (job.triggerID === "" && job.isValidJob()) {
                this.createTrigger(job, job.getMinutesToNextRun(runAsap) * 60);
            }
        }
    }

    /**
     * 
     * @param {CacheJob} job 
     * @param {Number} startAfterSeconds 
     * @returns {void}
     */
    createTrigger(job, startAfterSeconds) {
        if (job.timeout) {
            startAfterSeconds = 15;
        }

        const newTriggerID = ScriptApp
            .newTrigger('CacheFinanceTrigger')
            .timeBased()
            .after(startAfterSeconds * 1000)
            .create();
        job.triggerID = newTriggerID.getUniqueId();

        if (job.timeout) {
            const savedPartialResults = JSON.stringify(job.jobData);
            const key = CacheFinanceUtils.makeCacheKey("RESUMEJOB", job.triggerID);
            const shortCache = CacheService.getScriptCache();
            shortCache.put(key, savedPartialResults, 300);
            Logger.log(`Saving PARTIAL RESULTS.  Key=${key}. Items=${job.jobData.length}`);
        }

        this.updateFinanceCacheLegend();
        Logger.log(`New Trigger Created. ID=${job.triggerID}. Attrib=${job.attribute}. Start in ${startAfterSeconds} seconds.`);
    }

    /**
     * Update job status data to sheet.
     * @returns {void}
     */
    updateFinanceCacheLegend() {
        const legend = [];

        for (const job of this.jobs) {
            legend.push(job.toArray());
        }

        Logger.log(`Updating LEGEND.${legend}`);

        const sheetNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(CACHE_LEGEND);
        sheetNamedRange.setValues(legend);
    }

    /**
     * 
     * @param {GoogleAppsScript.Events.TimeDriven} e 
     * @returns {CacheJob}
     */
    getMyJob(e) {
        if (typeof e === 'undefined') {
            Logger.log("Trigger ID unknown.  Function not started as a TimeDriven event");
            return null;
        }

        let myJob = null;
        const runningScriptID = e.triggerUid;

        for (const job of this.jobs) {
            if (job.triggerID === runningScriptID) {
                Logger.log(`Found My Job ID=${job.triggerID}`);
                myJob = job;
                break;
            }
        }

        if (myJob === null) {
            Logger.log(`This JOB ${runningScriptID} not found in legend.`);
        }
        else {
            //  Is this a continuation job?
            const key = CacheFinanceUtils.makeCacheKey("RESUMEJOB", myJob.triggerID);
            const shortCache = CacheService.getScriptCache();
            const resumeData = shortCache.get(key);
            if (resumeData !== null) {
                myJob.wasRestarted = true;
                myJob.jobData = JSON.parse(resumeData);
                Logger.log(`Resuming Job: ${key}. Items added=${myJob.jobData.length}`);
            }
            else {
                Logger.log(`New Job.  Key=${key}`);
            }
        }

        return myJob;
    }
}

class CacheJob {
    /**
     * 
     * @param {any[]} jobParmameters 
     */
    constructor(jobParmameters) {
        this.DAYNAMES = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];

        const CACHE_SETTINGS_SYMBOL = 0;
        const CACHE_SETTINGS_ATTRIBUTE = 1;
        const CACHE_SETTINGS_OUTPUT = 2;
        const CACHE_SETTINGS_DEFAULT = 3;
        const CACHE_SETTINGS_REFRESH = 4;
        const CACHE_SETTINGS_HOURS = 5;
        const CACHE_SETTINGS_DAYS = 6;
        const CACHE_SETTINGS_ID = 7;

        this.symbolRange = jobParmameters[CACHE_SETTINGS_SYMBOL];
        this.attribute = jobParmameters[CACHE_SETTINGS_ATTRIBUTE];
        this.outputRange = jobParmameters[CACHE_SETTINGS_OUTPUT];
        this.defaultRange = jobParmameters[CACHE_SETTINGS_DEFAULT];
        this.refreshMinutes = jobParmameters[CACHE_SETTINGS_REFRESH];
        this.hours = jobParmameters[CACHE_SETTINGS_HOURS].toString();
        this.days = jobParmameters[CACHE_SETTINGS_DAYS].toString();
        this.triggerID = jobParmameters[CACHE_SETTINGS_ID];
        this.dayNumbers = [];
        this.hourNumbers = [];
        this.jobData = null;
        this.wasRestarted = false;
        this.maxRunSeconds = 0;
        this.startJobTime = null;

        Logger.log(`Job Settings:  symbols=${this.symbolRange}. attribute=${this.attribute}. Out=${this.outputRange}. In=${this.defaultRange}. Minutes=${this.refreshMinutes}. Hours=${this.hours}. Days=${this.days}`);

        if (typeof this.hours === 'string')
            this.hours = this.hours.trim().toUpperCase();

        if (typeof this.days === 'string')
            this.days = this.days.trim().toUpperCase();

        this.extractRunDaysOfWeek();
        this.extractHoursOfDay();
    }

    /**
     * 
     * @returns {Boolean}
     */
    isValidJob() {
        return this.symbolRange !== '' && this.outputRange !== '' && this.attribute !== '' && this.dayNumbers.length > 0 && this.hourNumbers.length > 0;
    }

    /**
     * 
     * @param {any[][]} data 
     */
    save(data) {
        this.jobData = data;
    }

    /**
     * 
     * @param {Boolean} value 
     */
    timedOut(value) {
        this.timeout = value;
    }

    /**
     * 
     * @returns {Boolean}
     */
    isRestarting() {
        return this.wasRestarted;
    }

    /**
     * @param {Number} maxRunSeconds
     * @returns {any[]}
     */
    startJobRunTimer(maxRunSeconds) {
        let data = [];
        if (this.isRestarting()) {
            data = this.jobData;
        }

        this.timedOut(false);
        this.maxRunSeconds = maxRunSeconds;
        this.startJobTime = new Date().getTime();

        return data;
    }

    /**
     * 
     * @returns {Boolean}
     */
    hasExceededJobRunTime() {
        const elapsed = (new Date().getTime() - this.startJobTime) / 1000;
        if (elapsed > this.maxRunSeconds) {
            Logger.log("Max. Job Time reached.");
            this.timedOut(true);
            return true;
        }

        return false;
    }

    /**
     * Extract days that can be run from job legend into our job info.
     * @returns {void}
     */
    extractRunDaysOfWeek() {
        this.dayNumbers = [];

        if (this.days === '') {
            this.days = '*';
        }

        for (let i = 0; i < 7; i++)
            this.dayNumbers[i] = false;

        if (this.days === '*') {
            this.days = this.DAYNAMES.join(",");
        }

        if (this.extractListDaysOfWeek()) {
            return;
        }

        if (this.extractRangeDaysOfWeek()) {
            return;
        }

        const dayNameIndex = this.DAYNAMES.indexOf(this.days);
        if (dayNameIndex !== -1) {
            this.dayNumbers[dayNameIndex] = true;
            return;
        }

        const singleItem = parseInt(this.days, 10);
        if (!isNaN(singleItem)) {
            this.dayNumbers[singleItem] = true;
            return;
        }

        Logger.log(`Job does not contain any days of the week to run. ${this.days}`);
    }

    /**
     * 
     * @returns {Boolean}
     */
    extractListDaysOfWeek() {
        if (this.days.indexOf(",") === -1) {
            return false;
        }

        let listValues = this.days.split(",");
        listValues = listValues.map(p => p.trim());

        for (const day of listValues) {
            let dayNameIndex = this.DAYNAMES.indexOf(day);
            if (dayNameIndex === -1) {
                dayNameIndex = parseInt(day, 10);
                if (isNaN(dayNameIndex))
                    dayNameIndex = -1;

                if (dayNameIndex < 0 || dayNameIndex > 6)
                    dayNameIndex = -1;
            }

            if (dayNameIndex !== -1) {
                this.dayNumbers[dayNameIndex] = true;
            }
        }

        return true;
    }

    /**
     * 
     * @returns {Boolean}
     */
    extractRangeDaysOfWeek() {
        if (this.days.indexOf("-") === -1) {
            return false;
        }

        let rangeValues = this.days.split("-");
        rangeValues = rangeValues.map(p => p.trim());

        let startDayNameIndex = this.DAYNAMES.indexOf(rangeValues[0]);
        let endDayNameIndex = this.DAYNAMES.indexOf(rangeValues[1]);
        if (startDayNameIndex === -1 || endDayNameIndex === -1) {
            startDayNameIndex = parseInt(rangeValues[0], 10);
            endDayNameIndex = parseInt(rangeValues[1], 10);

            if (isNaN(startDayNameIndex) || isNaN(endDayNameIndex)) {
                return false;
            }
        }

        let count = 0;
        for (let i = startDayNameIndex; count < 7; count++) {
            this.dayNumbers[i] = true;

            if (i === endDayNameIndex)
                break;

            i++;
            if (i > 6)
                i = 0;
        }

        return true;
    }

    /**
     * Parse HOURS set in legend into 24 hour true/false array.
     * @returns {void}
     */
    extractHoursOfDay() {
        this.hourNumbers = [];

        if (this.hours === '') {
            this.hours = '*';
        }

        for (let i = 0; i < 24; i++)
            this.hourNumbers[i] = this.hours === '*';

        if (this.hours === '*')
            return;

        if (this.extractListHoursOfDay())
            return;


        if (this.extractRangeHoursOfDay())
            return;

        const singleItem = parseInt(this.hours, 10);
        if (!isNaN(singleItem)) {
            this.hourNumbers[singleItem] = true;
            return;
        }

        Logger.log(`This job does not contain any valid hours to run. ${this.hours}`);
    }

    /**
     * 
     * @returns {Boolean}
     */
    extractListHoursOfDay() {
        if (this.hours.indexOf(",") === -1) {
            return false;
        }

        let listValues = this.hours.split(",");
        listValues = listValues.map(p => p.trim());

        for (const hr of listValues) {
            let hourIndex = parseInt(hr, 10);
            if (isNaN(hourIndex))
                hourIndex = -1;

            if (hourIndex < 0 || hourIndex > 23)
                hourIndex = -1;


            if (hourIndex !== -1) {
                this.hourNumbers[hourIndex] = true;
            }
        }

        return true;
    }

    /**
     * 
     * @returns {Boolean}
     */
    extractRangeHoursOfDay() {
        if (this.hours.indexOf("-") === -1) {
            return false;
        }

        let rangeValues = this.hours.split("-");
        rangeValues = rangeValues.map(p => p.trim());

        const startHourIndex = parseInt(rangeValues[0], 10);
        const endHourIndex = parseInt(rangeValues[1], 10);

        if (isNaN(startHourIndex) || isNaN(endHourIndex)) {
            return false;
        }

        if (startHourIndex < 0 || startHourIndex > 23 ||
            endHourIndex < 0 || endHourIndex > 23) {
            return false;
        }

        let count = 0;
        for (let i = startHourIndex; count < 24; count++) {
            this.hourNumbers[i] = true;

            if (i === endHourIndex)
                break;

            i++;
            if (i > 23)
                i = 0;
        }

        return true;
    }

    /**
     * 
     * @param {Boolean} runAsap 
     * @returns {Number}
     */
    getMinutesToNextRun(runAsap = false) {
        let minutes = this.refreshMinutes < 1 || runAsap ? 1 : this.refreshMinutes;

        // Get current date
        const startDateTime = new Date();
        let daysSinceStart = 0;
        const date = new Date();
        date.setMinutes(date.getMinutes() + minutes);

        //  The next run is within our window of opportunity.
        while (!this.canRunJob(date) && daysSinceStart <= 7) {
            if (!this.canRunJobForDayOfWeek(date) ||
                (this.canRunJobForDayOfWeek(date) && !this.canRunJobLaterToday(date))) {
                date.setDate(date.getDate() + 1);
                date.setHours(0);
                date.setMinutes(0)
            }
            else {
                date.setHours(this.getNextHourToRun(date));
                date.setMinutes(0);
            }

            daysSinceStart = (date.getTime() - startDateTime.getTime()) / (1000 * 3600 * 24);
        }

        if (daysSinceStart > 7) {
            throw new Error("Error finding next TRIGGER time for job.");
        }

        minutes = (date.getTime() - startDateTime.getTime()) / 60000;

        return minutes;
    }

    /**
     * 
     * @param {Date} startDate 
     * @returns {boolean}
     */
    canRunJob(startDate = new Date()) {
        return this.canRunJobForDayOfWeek(startDate) && this.canRunJobForTime(startDate);
    }

    /**
     * 
     * @param {Date} date 
     * @returns {Boolean}
     */
    canRunJobForDayOfWeek(date) {
        const dow = date.getDay();
        const status = this.dayNumbers[dow];

        return status;
    }

    /**
     * 
     * @param {Date} date 
     * @returns {Boolean}
     */
    canRunJobForTime(date) {
        const hour = date.getHours();
        const status = this.hourNumbers[hour];

        return status;
    }

    /**
     * 
     * @param {Date} date 
     * @returns {Boolean}
     */
    canRunJobLaterToday(date) {
        return this.getNextHourToRun(date) !== -1;
    }

    /**
     * 
     * @param {Date} date 
     * @returns {Number}
     */
    getNextHourToRun(date) {
        const hr = date.getHours();
        for (let i = hr; i < 24; i++) {
            if (this.hourNumbers[i])
                return i;
        }
        return -1;
    }

    /**
     * 
     * @returns {any[]}
     */
    toArray() {
        const row = [];
        row.push(this.symbolRange);
        row.push(this.attribute);
        row.push(this.outputRange);
        row.push(this.defaultRange);
        row.push(this.refreshMinutes);
        row.push(this.hours);
        row.push(this.days);
        row.push(this.triggerID);

        return row;
    }
}


class CacheTrigger {
    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {Boolean} - true - DELETE JOB and RE-CREATE.
     */
    static runJob(jobSettings) {
        if (jobSettings === null) {
            Logger.log("Job settings not found.");
            return false;
        }

        if (!jobSettings.canRunJob()) {
            Logger.log("* * *   Not time to run JOB   * * *");
            return true;
        }

        Logger.log("Starting JOB.  Updating FINANCE info.");

        CacheTrigger.getJobData(jobSettings);
        CacheTrigger.writeResults(jobSettings);

        return true;
    }

    /**
     * 
     * @param {CacheJob} jobSettings - results also returned in object.
     * @returns {void}
     */
    static getJobData(jobSettings) {
        const MAX_RUN_SECONDS = 300;
        const attribute = jobSettings.attribute;
        const symbols = CacheTrigger.getSymbols(jobSettings);
        const defaultData = CacheTrigger.getDefaultData(jobSettings);
        const bulkDataCache = CacheFinanceUtils.bulkShortCacheGet(symbols, attribute);

        Logger.log(`CacheTrigger: Symbols:${jobSettings.symbolRange}. Attribute:${jobSettings.attribute}. GoogleRange: ${jobSettings.defaultRange}`);

        if (CacheTrigger.isDefaultDataProvided(jobSettings) && symbols.length !== defaultData.length) {
            throw new Error(`Symbol Ranges and Google Values Ranges must be the same: ${jobSettings.symbolRange}. Len=${symbols.length}.  GoogleValues: ${jobSettings.defaultRange}. Len=${defaultData.length}`);
        }

        //  Start job timer and load previously loaded (and not finished) job data.
        const data = jobSettings.startJobRunTimer(MAX_RUN_SECONDS);
        const startingSymbol = data.length;

        for (let i = startingSymbol; i < symbols.length; i++) {
            if (jobSettings.hasExceededJobRunTime()) {
                break;
            }

            let value = null;
            if (CacheTrigger.isDefaultDataProvided(jobSettings)) {
                value = CacheFinance.getFinanceData(symbols[i][0], attribute, defaultData[i][0], bulkDataCache[i]);
            }
            else {
                value = CacheFinance.getFinanceData(symbols[i][0], attribute, GOOGLEFINANCE_PARAM_NOT_USED);
            }

            data.push([value]);
        }

        jobSettings.save(data);
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {any[][]}
     */
    static getSymbols(jobSettings) {
        const symbolsNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.symbolRange);
        if (symbolsNamedRange === null) {
            throw new Error(`Failed to read data from SYMBOLS range =${jobSettings.symbolRange}`);
        }

        return symbolsNamedRange.getValues();
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {any[][]}
     */
    static getDefaultData(jobSettings) {
        let defaultData = [];

        if (jobSettings.defaultRange !== "") {
            const symbolsNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.defaultRange);
            if (symbolsNamedRange === null) {
                throw new Error(`Failed to read data from DEFAULT DATA range =${jobSettings.defaultRange}`);
            }

            defaultData = symbolsNamedRange.getValues();
        }

        return defaultData;
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {Boolean}
     */
    static isDefaultDataProvided(jobSettings) {
        return jobSettings.defaultRange !== "";
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {Boolean}
     */
    static writeResults(jobSettings) {
        if (jobSettings.jobData === null || jobSettings.timeout)
            return false;

        Logger.log(`writeCacheResults:  START.  Data Len=${jobSettings.jobData.length}`);

        const outputNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.outputRange);

        if (outputNamedRange === null) {
            return false;
        }

        try {
            outputNamedRange.setValues(jobSettings.jobData);
        }
        catch (ex) {
            Logger.log(`Updating output range FAILED.  ${ex.toString()}`);
            return false;
        }

        Logger.log(`writeCacheResults:  END: ${jobSettings.outputRange}`);
        return true;
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
     * 
     * @param {Object} propertyDataObject 
     * @param {Number} daysToHold 
     */
    putAll(propertyDataObject, daysToHold = 1) {
        const keys = Object.keys(propertyDataObject);
        keys.forEach(key => this.put(key, propertyDataObject[key], daysToHold));
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
        const allProperties = PropertiesService.getScriptProperties().getProperties();

        //  Removing properties is very slow, so remove only 1 at a time.  This is enough as this function is called frequently.
        ScriptSettings.expire(false, 1, allProperties);

        for (const key of cacheKeys) {
            const myData = allProperties[key];

            if (typeof myData === 'undefined') {
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
     * 
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
        const MAX_TIME_FOR_FETCH_Ms = 25000;        //  Custom function times out at 30 seconds, so we need to limit.
        const bestStockSites = FinanceWebsiteSearch.readBestStockWebsites();
        const siteURLs = FinanceWebsiteSearch.getAllStockWebSiteFunctions(symbols, attribute, bestStockSites);
        
        let batch = 1;
        const startTime = Date.now();
        let missingStockData = [...siteURLs];
        while (missingStockData.length > 0 && (Date.now() - startTime) < MAX_TIME_FOR_FETCH_Ms) { 
            const URLs = FinanceWebsiteSearch.getNextUrlBatch(missingStockData);

            Logger.log(`Batch=${batch}. URLs${URLs}`);
            const responses = FinanceWebsiteSearch.bulkSiteFetch(URLs);
            const elapsedTime = Date.now() - startTime;
            Logger.log(`Batch=${batch}. Responses=${responses.length}. Total Elapsed=${elapsedTime}`);
            batch++;

            FinanceWebsiteSearch.updateStockResults(missingStockData, URLs, responses, attribute, bestStockSites);
            
            missingStockData = missingStockData.filter(stock => ! stock.stockAttributes.isAttributeSet(attribute) && ! stock.isSitesDone())
        }

        //  TODO:  If separate CACHEFINANCES() run at the same time, the last process to finish will overwrite any new results
        //         from the other runs.  This is not critical, since it is ONLY  used to improve the ordering of sites to call AND
        //         over time as the processes run on their own, the data will be corrected.
        FinanceWebsiteSearch.writeBestStockWebsites(bestStockSites);
        
        return siteURLs.map(stock => stock.stockAttributes);
    }

    /**
     * 
     * @param {StockWebURL[]} missingStockData 
     * @returns {String[]}
     */
    static getNextUrlBatch(missingStockData) {
        const MAX_FETCHALL_BATCH_SIZE = 50;

        let URLs = missingStockData.map(url => url.getURL()).filter(url => url !== null && url !== '');
        if (URLs.length > MAX_FETCHALL_BATCH_SIZE) {
            URLs = URLs.slice(0, MAX_FETCHALL_BATCH_SIZE);
        }

        return URLs;
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
            const matchingSites = missingStockData.filter(site => site.getURL() === URLs[i]);
            matchingSites.map(site => site.parseResponse(responses[i], attribute));
            matchingSites.map(site => site.updateBestSites(bestStockSites, attribute));
            matchingSites.map(site => site.skipToNextSite());
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
        const siteInfo = new FinanceWebSites();
        const siteList = siteInfo.get();

        //  Getting this is slow, so save and use later.
        for (const site of siteList) {
            apiMap.set(site.siteName, site.siteObject.getApiKey())
        }

        for (const symbol of symbols) {
            const stockURL = new StockWebURL(symbol);
            const bestSite = bestStockSites[CacheFinanceUtils.makeCacheKey(symbol, attribute)];

            for (const site of siteList) {
                stockURL.addSiteURL(site.siteName, site.siteName === bestSite, site.siteObject.getURL(symbol, attribute, apiMap.get(site.siteName)), site.siteObject.parseResponse);
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
        longCache.put("CACHE_WEBSITES", siteObject, 365);    
    }


    /**
     * 
     * @param {String[]} URLs 
     * @returns {String[]}
     */
    static bulkSiteFetch(URLs) {
        const filteredURLs = URLs.filter(url => url.trim() !== '');
        const fetchURLs = filteredURLs.map(url => {
            //  skipcq:  JS-0240
            return {
                'url': url,
                'method': 'get',
                'muteHttpExceptions': true
            }
        });

        let dataSet = [];
        try {
            const rawSiteData = UrlFetchApp.fetchAll(fetchURLs);
            dataSet = rawSiteData.map(response => response.getContentText());
        }
        catch (ex) {
            return dataSet;
        };

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

class StockWebURL {
    constructor (symbol) {
        this.symbol = symbol;
        this.siteName = [];
        this.siteURL = [];
        this.bestSites = [];
        this.parseFunction = [];
        /** @type {StockAttributes} */
        this.stockAttributes = new StockAttributes();
        this.siteIterator = 0;
    }

    /**
     * 
     * @param {String} siteName 
     * @param {String} URL 
     * @param {Object} parseResponseFunction 
     * @returns 
     */
    addSiteURL(siteName, bestSite, URL, parseResponseFunction) {
        if (URL.trim() === '') {
            return;
        } 

        if (bestSite) {
            this.siteName.unshift(siteName);
            this.siteURL.unshift(URL);
            this.parseFunction.unshift(parseResponseFunction);
            this.bestSites.unshift(true);
        }
        else {
            this.siteName.push(siteName);
            this.siteURL.push(URL);
            this.parseFunction.push(parseResponseFunction);
            this.bestSites.push(false);
        }
    }

    /**
     * Returns next website URL to be used.
     * @returns {String}
     */
    getURL() {
        return  this.siteIterator < this.siteURL.length ? this.siteURL[this.siteIterator] : null;  
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
        this.bestSites[this.siteIterator] = (this.stockAttributes === null || ! this.stockAttributes.isAttributeSet(attribute)) ? false : true;

        return this.stockAttributes;
    }

    /**
     * 
     * @param {Object} bestStockSites 
     * @param {String} attribute 
     */
    updateBestSites(bestStockSites, attribute) {
        const key = CacheFinanceUtils.makeCacheKey(this.symbol, attribute);
        bestStockSites[key] = (! this?.stockAttributes.isAttributeSet(attribute)) ? "" : this.siteName[this.siteIterator];
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
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "NYSEARCA:SHYG");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "TSE:MEG");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "TSE:RY");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "NASDAQ:MSFT");
        
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "NYSEARCA:SHYG");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:ZTL");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:DFN-A");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:DFN-A", "ALL", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "NASDAQ:MSFT", "ALL", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:RY", "ALL", "STOCK");

        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NYSEARCA:SHYG");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:FTN-A");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:RY");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NASDAQ:MSFT");

        const plan = new FinanceWebsiteSearch();
        //  Make fresh lookup plans.
        FinanceWebsiteSearch.deleteLookupPlan("TSE:RY");
        FinanceWebsiteSearch.deleteLookupPlan("NASDAQ:BNDX");
        FinanceWebsiteSearch.deleteLookupPlan("TSE:ZTL");

        plan.getLookupPlan("TSE:RY", "");
        plan.getLookupPlan("NASDAQ:BNDX", "");
        plan.getLookupPlan("TSE:ZTL", "");

        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "TSE:RY", "PRICE");
        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "NASDAQ:BNDX", "PRICE");
        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "NASDAQ:BNDX", "NAME");
        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "NASDAQ:BNDX", "YIELDPCT");
        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "TSE:ZTL", "PRICE");

        CacheFinance.deleteFromCache("TSE:RY", "PRICE");
        this.cacheTestRun.run("CACHEFINANCE - not cached", CACHEFINANCE, "TSE:RY", "PRICE", "##NotSet##");
        this.cacheTestRun.run("CACHEFINANCE - cached", CACHEFINANCE, "TSE:RY", "PRICE", "##NotSet##");


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

            if (! (data instanceof StockAttributes)) {
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
        catch(ex) {
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

            row.push(testRun.serviceName);
            row.push(testRun.symbol);
            row.push(testRun.status);
            row.push(testRun.stockAttributes.stockPrice);
            row.push(testRun.stockAttributes.yieldPct);
            row.push(testRun.stockAttributes.stockName);
            row.push(testRun._attributeLookup);
            row.push(testRun.typeLookup);
            row.push(testRun.runTime);

            resultTable.push(row);
        }

        return resultTable;
    }
}

/**
 * @classdesc Individual test results and tracking.
 */
class CacheFinanceTestStatus {
    constructor (serviceName="", symbol="") {
        this._serviceName = serviceName;
        this._symbol = symbol;
        this._stockAttributes = new StockAttributes();
        this._startTime = Date.now()
        this._typeLookup = "";
        this._attributeLookup  = "";
        this._runTime = 0;
    }

    get serviceName() {
        return this._serviceName;
    }
    get symbol () {
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
     * All finance Website Objects must implement the method "getInfo(symbol)" and getPropertyValue(key)
     * The getInfo() method must return an instance of "StockAttributes"
     * The getPropertyValue() is used to query any possible properties known (and unknown future) about the website.
     */
    constructor() {
        this.siteList = [
            new FinanceWebSite("FinnHub", FinnHub),           
            new FinanceWebSite("TDEtf", TdMarketsEtf),
            new FinanceWebSite("TDStock", TdMarketsStock),
            new FinanceWebSite("Globe", GlobeAndMail),
            new FinanceWebSite("Yahoo", YahooFinance),
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

        return (typeof siteInfo === 'undefined') ? null : siteInfo;
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
                countryCode = "us";
                break;
            case "CVE":
            case "TSE":
            case "TSX":
            case "TSXV":
                countryCode = "ca";
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
                return this.stockPrice !== null && this.stockPrice !== 0;

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
 * @classdesc TD Markets lookup by ETF symbol
 */
class TdMarketsEtf {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        return TdMarketResearch.getInfo(symbol, attribute, "ETF");
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
        return TdMarketResearch.getURL(symbol, "ETF");
    }

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
        return TdMarketResearch.parseResponse(html);
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
    }
}

/**
 * @classdesc TD Markets lookup by STOCK symbol
 */
class TdMarketsStock {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        return TdMarketResearch.getInfo(symbol, attribute, "STOCK");
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
        return TdMarketResearch.getURL(symbol, "STOCK");
    }

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
        return TdMarketResearch.parseResponse(html);
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
    }
}

/**
 * @classdesc Base class to Lookup for TD website.
 */
class TdMarketResearch {
    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @param {String} type 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, _attribute, type = "ETF") {
        const URL = TdMarketResearch.getURL(symbol, type);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }
        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        return TdMarketResearch.parseResponse(html);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} type 
     * @returns {String}
     */
    static getURL(symbol, type = "ETF") {
        let URL = null;
        if (type === "ETF")
            URL = `https://marketsandresearch.td.com/tdwca/Public/ETFsProfile/Summary/${FinanceWebSites.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;
        else
            URL = `https://marketsandresearch.td.com/tdwca/Public/Stocks/Overview/${FinanceWebSites.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;

        return URL;
    }

    /**
     * 
     * @param {String} html 
     * @returns {StockAttributes}
     */
    static parseResponse(html) {
        const data = new StockAttributes();

        //  Get the dividend yield.

        let parts = html.match(/Dividend Yield<\/th><td class="last">(\d{0,4}\.?\d{0,4})%/);
        if (parts === null) {
            parts = html.match(/Dividend Yield<\/div>.*?cell-container contains">(\d{0,4}\.?\d{0,4})%/);
        }
        if (parts !== null && parts.length === 2) {
            const tempPct = parts[1];

            const parsedValue = parseFloat(tempPct) / 100;

            if (!isNaN(parsedValue)) {
                data.yieldPct = parsedValue;
            }
        }

        //  Get the name.
        parts = html.match(/.<span class="issueName">(.*?)<\//);
        if (parts !== null && parts.length === 2) {
            data.stockName = parts[1];
        }

        //  Get the price.
        parts = html.match(/.LAST PRICE<\/span><div><span>(\d{0,4}\.?\d{0,4})</);
        if (parts !== null && parts.length === 2) {

            const parsedValue = parseFloat(parts[1]);
            if (!isNaN(parsedValue)) {
                data.stockPrice = parsedValue;
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
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const parts = symbol.split(":");
            symbol = parts[1];
        }

        const dash = symbol.indexOf("-");
        if (dash >= 0) {
            symbol = symbol.replace("-", ".PR.");
        }

        return symbol;
    }
}

/**
 * @classdesc Lookup for Yahoo site.
 */
class YahooFinance {
    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        const URL = YahooFinance.getURL(symbol);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
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
        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (countryCode !== "us") {
            return "";
        }

        return `https://finance.yahoo.com/quote/${YahooFinance.getTicker(symbol)}`;
    }

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

        let dividendPercent = html.match(/"DIVIDEND_AND_YIELD-value">\d*\.\d*\s\((\d*\.\d*)%\)/);
        if (dividendPercent === null) {
            dividendPercent = html.match(/TD_YIELD-value">(\d*\.\d*)%/);
        }

        if (dividendPercent !== null && dividendPercent.length === 2) {
            const tempPct = dividendPercent[1];
            Logger.log(`Yahoo. Stock=${symbol}. PERCENT=${tempPct}`);

            data.yieldPct = parseFloat(tempPct) / 100;

            if (isNaN(data.yieldPct)) {
                data.yieldPct = null;
            }
        }

        const baseSymbol = YahooFinance.getTicker(symbol);
        const re = new RegExp(`data-symbol="${baseSymbol}".+?"regularMarketPrice".+?value="(\\d*\\.?\\d*)?"`);

        const priceMatch = html.match(re);

        if (priceMatch !== null && priceMatch.length === 2) {
            const tempPrice = priceMatch[1];
            Logger.log(`Yahoo. Stock=${symbol}.PRICE=${tempPrice}`);

            data.stockPrice = parseFloat(tempPrice);

            if (isNaN(data.stockPrice)) {
                data.stockPrice = null;
            }
        }

        return data;
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
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

        }
        return modifiedSymbol;
    }
}

/**
 * @classdesc Lookup for Globe and Mail website.
 */
class GlobeAndMail {
    /**
     * Only gets dividend yield.
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        const URL = GlobeAndMail.getURL(symbol);

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }

        return GlobeAndMail.parseResponse(html);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
        return `https://www.theglobeandmail.com/investing/markets/stocks/${GlobeAndMail.getTicker(symbol)}`;
    }

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

        if (parts === null)
            parts = html.match(/.name=\\"dividendYieldTrailing\\".*?value=\\"(\d{0,4}\.?\d{0,4})%/);

        if (parts !== null && parts.length === 2) {
            const tempPct = parts[1];

            const parsedValue = parseFloat(tempPct) / 100;

            if (!isNaN(parsedValue)) {
                data.yieldPct = parsedValue;
            }
        }

        //  Get the name.
        parts = html.match(/."symbolName":"(.*?)"/);
        if (parts !== null && parts.length === 2) {
            data.stockName = parts[1];
        }

        //  Get the price.
        parts = html.match(/."lastPrice":"(\d{0,4}\.?\d{0,4})"/);
        if (parts !== null && parts.length === 2) {

            const parsedValue = parseFloat(parts[1]);
            if (!isNaN(parsedValue)) {
                data.stockPrice = parsedValue;
            }
        }

        return data;
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
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
                    if (parts[1].indexOf(".") !== -1) {
                        symbol = parts[1].replace(".", "-");
                    }
                    else if (parts[1].indexOf("-") !== -1) {
                        const prefShare = parts[1].split("-");
                        symbol = `${prefShare[0]}-PR-${prefShare[1]}`;
                    }
                    symbol = `${symbol}-T`;
                    break;
                case "NYSEARCA":
                    symbol = `${parts[1]}-A`;
                    break;
                case "NASDAQ":
                    symbol = `${parts[1]}-Q`;
                    break;
                default:
                    symbol = '#N/A';
            }
        }

        return symbol;
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
            data = FinnHub.parseResponse(jsonStr);
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
    static getURL(symbol, attribute, API_KEY=null) {
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
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
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
        if (countryCode !== "us") {
            Logger.log(`AlphaVantage --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const URL = AlphaVantage.getURL(symbol, attribute, AlphaVantage.getApiKey());
        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();
            data = AlphaVantage.parseResponse(jsonStr);
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
    static getURL(symbol, attribute, API_KEY=null) {
        if (API_KEY === null) {
            return "";
        }

        if (attribute !== "PRICE") {
            return "";
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (countryCode !== "us") {
            return "";
        }

        return `https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=${FinanceWebSites.getBaseTicker(symbol)}&apikey=${API_KEY}`;
    }

    static getApiKey() {
        return FinanceWebSites.getApiKey("ALPHA_VANTAGE_API_KEY");    
    }

    /**
     * 
     * @param {String} jsonStr 
     * @returns {StockAttributes}
     */
    static parseResponse(jsonStr) {
        const data = new StockAttributes();

        Logger.log(`content=${jsonStr}`);
        try {
            const alphaVantageData = JSON.parse(jsonStr);
            data.stockPrice = alphaVantageData["Global Quote"]["05. price"];
            Logger.log(`Price=${data.stockPrice}`);
        }
        catch (ex) {
            Logger.log("AlphaVantage JSON Parse Error.");
        }

        return data;
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
    }
}

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

