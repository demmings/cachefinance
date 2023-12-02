
const GOOGLEFINANCE_PARAM_NOT_USED = "##NotSet##";

//  Function only used for testing in google sheets app script.
// skipcq: JS-0128
function testYieldPct() {
    const val = CACHEFINANCE("TSE:FTN-A", "yieldpct");        // skipcq: JS-0128
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
        const ss = new ScriptSettings();
        ss.expire(true);
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
     * @returns {any}
     */
    static getFinanceData(symbol, attribute, googleFinanceValue) {
        attribute = attribute.toUpperCase().trim();
        symbol = symbol.toUpperCase();
        const cacheKey = CacheFinance.makeCacheKey(symbol, attribute);

        //  This time GOOGLEFINANCE worked!!!
        if (googleFinanceValue !== GOOGLEFINANCE_PARAM_NOT_USED && googleFinanceValue !== "#N/A" && googleFinanceValue !== '#ERROR!') {
            //  We cache here longer because we would normally be getting data from Google.
            //  If GoogleFinance is failing, we need the data to be held longer since it
            //  it is getting from cache as an emergency backup.
            CacheFinance.saveFinanceValueToCache(cacheKey, googleFinanceValue, 21600);
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
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {String}
     */
    static makeCacheKey(symbol, attribute) {
        return `${attribute}|${symbol}`;
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
            CacheFinance.saveFinanceValueToCache(CacheFinance.makeCacheKey(symbol, "NAME"), stockAttributes.stockName, 1200);
        if (stockAttributes.isAttributeSet("PRICE"))
            CacheFinance.saveFinanceValueToCache(CacheFinance.makeCacheKey(symbol, "PRICE"), stockAttributes.stockPrice, 1200);
        if (stockAttributes.isAttributeSet("YIELDPCT"))
            CacheFinance.saveFinanceValueToCache(CacheFinance.makeCacheKey(symbol, "YIELDPCT"), stockAttributes.yieldPct, 1200);
    }

    /**
     * 
     * @param {String} key 
     * @param {any} financialData 
     * @param {Number} shortCacheSeconds 
     * @param {Number} longCacheDays
     * @returns {void}
     */
    static saveFinanceValueToCache(key, financialData, shortCacheSeconds = 1200, longCacheDays=7) {
        const shortCache = CacheService.getScriptCache();
        const longCache = new ScriptSettings();
        let start = new Date().getTime();

        const currentShortCacheValue = shortCache.get(key);
        if (currentShortCacheValue !== null && JSON.parse(currentShortCacheValue) === financialData) {
            Logger.log(`GoogleFinance VALUE.  No Change in SHORT Cache. ms=${new Date().getTime() - start}`);
            return;
        }

        if (currentShortCacheValue !== null) {
            Logger.log(`Short Cache Changed.  Old=${JSON.parse(currentShortCacheValue)} . New=${financialData}`);
        }
   
        //  If we normally get the price from Google, we want to cache for a longer
        //  time because the only time we need a price for this particular stock
        //  is when GOOGLEFINANCE fails.
        start = new Date().getTime();
        shortCache.put(key, JSON.stringify(financialData), shortCacheSeconds);
        const shortMs = new Date().getTime() - start;
       
        //  For emergency cases when GOOGLEFINANCE is down long term...
        start = new Date().getTime();
        longCache.put(key, financialData, longCacheDays);
        const longMs = new Date().getTime() - start;

        Logger.log(`SET GoogleFinance VALUE Long/Short Cache. Key=${key}.  Value=${financialData}. Short ms=${shortMs}. Long ms=${longMs}`);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     */
    static deleteFromCache(symbol, attribute) {
        const key = CacheFinance.makeCacheKey(symbol, attribute);
        
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
     * Removes script settings that have expired.
     * @param {Boolean} deleteAll - true - removes ALL script settings regardless of expiry time.
     */
    expire(deleteAll) {
        const allKeys = this.scriptProperties.getKeys();

        for (const key of allKeys) {
            const myData = this.scriptProperties.getProperty(key);

            if (myData !== null) {
                let propertyValue = null;
                try {
                    propertyValue = JSON.parse(myData);
                }
                catch (e) {
                    Logger.log(`Script property data is not JSON. key=${key}`);
                    continue;
                }

                const propertyOfThisApplication = propertyValue !== null && propertyValue.expiry !== undefined;

                if (propertyOfThisApplication && (PropertyData.isExpired(propertyValue) || deleteAll)) {
                    this.scriptProperties.deleteProperty(key);
                    Logger.log(`Removing expired SCRIPT PROPERTY: key=${key}`);
                }
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

/** Converts data into JSON for getting/setting in ScriptSettings. */
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
            if (!PropertyData.isExpired(obj)) {
                value = JSON.parse(obj.myData);
            }
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

        this.priceSites.sort((a,b) => a.timeMs - b.timeMs);
        this.nameSites.sort((a,b) => a.timeMs - b.timeMs);
        this.yieldSites.sort((a,b) => a.timeMs - b.timeMs);
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
            catch(ex) {
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
            new FinanceWebSite("AlphaVantage", AlphaVantage),
            new FinanceWebSite("TDEtf", TdMarketsEtf),
            new FinanceWebSite("TDStock", TdMarketsStock),
            new FinanceWebSite("Yahoo", YahooFinance),
            new FinanceWebSite("Globe", GlobeAndMail)
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
        this.siteObject = webSiteJsClass;
    }

    set siteName(siteName) {
        this._siteName = siteName;
    }
    get siteName() {
        return this._siteName;
    }

    set webSiteClass(siteObject) {
        this._siteObject = siteObject;
    }
    get webSiteClass() {
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
                return (this.stockPrice === null) ? 0 : this.stockPrice;

            case "YIELDPCT":
                return (this.yieldPct === null) ? 0 : this.yieldPct;

            case "NAME":
                return (this.stockName === null) ? "" : this.stockName;

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
                return this.stockPrice !== null;

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
     * @param {String} attribute
     * @param {String} type 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute, type = "ETF") {
        const data = new StockAttributes();

        let URL = null;
        if (type === "ETF")
            URL = `https://marketsandresearch.td.com/tdwca/Public/ETFsProfile/Summary/${FinanceWebSites.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;
        else
            URL = `https://marketsandresearch.td.com/tdwca/Public/Stocks/Overview/${FinanceWebSites.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return data;
        }
        Logger.log(`getInfo:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

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
        const data = new StockAttributes();

        const URL = `https://finance.yahoo.com/quote/${YahooFinance.getTicker(symbol)}`;

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return data;
        }

        Logger.log(`getInfo:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        let dividendPercent = html.match(/"DIVIDEND_AND_YIELD-value">\d*\.\d*\s\((\d*\.\d*)%\)/);
        if (dividendPercent === null) {
            dividendPercent = html.match(/TD_YIELD-value">(\d*\.\d*)%/);
        }

        if (dividendPercent !== null && dividendPercent.length === 2) {
            const tempPct = dividendPercent[1];
            Logger.log(`PERCENT=${tempPct}`);

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
            Logger.log(`PRICE=${tempPrice}`);

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
        const data = new StockAttributes();
        const URL = `https://www.theglobeandmail.com/investing/markets/stocks/${GlobeAndMail.getTicker(symbol)}`;

        Logger.log(`getInfo:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return data;
        }

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
        const data = new StockAttributes();
        const API_KEY = FinanceWebSites.getApiKey("FINNHUB_API_KEY");

        if (API_KEY === null) {
            Logger.log("No FinnHub API Key.");
            return data;
        }

        if (attribute !== "PRICE") {
            Logger.log(`Finnhub.  Only PRICE is supported: ${symbol}, ${attribute}`);
            return data;
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (countryCode !== "us") {
            Logger.log(`FinnHub --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const URL = `https://finnhub.io/api/v1/quote?symbol=${FinanceWebSites.getBaseTicker(symbol)}&token=${API_KEY}`;
        Logger.log(`getInfo:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();

            const hubData = JSON.parse(jsonStr);
            data.stockPrice = hubData.c;
            Logger.log(hubData);
        }
        catch (ex) {
            return data;
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

class AlphaVantage {

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns 
     */
    static getInfo(symbol, attribute = "PRICE") {
        const data = new StockAttributes();
        const API_KEY = FinanceWebSites.getApiKey("ALPHA_VANTAGE_API_KEY");

        if (API_KEY === null) {
            Logger.log("No AlphaVantage API Key.");
            return data;
        }

        if (attribute !== "PRICE") {
            Logger.log(`AlphaVantage.  Only PRICE is supported: ${symbol}, ${attribute}`);
            return data;
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (countryCode !== "us") {
            Logger.log(`AlphaVantage --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const URL = `https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=${FinanceWebSites.getBaseTicker(symbol)}&apikey=${API_KEY}`;
        Logger.log(`getInfo:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();

            const alphaVantageData = JSON.parse(jsonStr);
            data.stockPrice = alphaVantageData["Global Quote"]["05. price"];
            Logger.log(`content=${jsonStr}`);
            Logger.log(`Price=${data.stockPrice}`);
        }
        catch (ex) {
            return data;
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

