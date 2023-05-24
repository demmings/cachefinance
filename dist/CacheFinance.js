const GOOGLEFINANCE_PARAM_NOT_USED = "##NotSet##";

//  Named range in sheet with CacheFinance configurations.
const CACHE_LEGEND = "CACHEFINANCE";

/**
 * Replacement function to GOOGLEFINANCE for stock symbols not recognized by google.
 * @param {string} symbol 
 * @param {string} attribute - ["price", "yieldpct", "name"] - "TEST" - returns test results from 3rd party sites.
 * @param {any} googleFinanceValue - Optional.  Use GOOGLEFINANCE() to get value, if '#N/A' will read cache.
 * @returns {any}
 * @customfunction
 */
function CACHEFINANCE(symbol, attribute = "price", googleFinanceValue = GOOGLEFINANCE_PARAM_NOT_USED) {
    Logger.log(`CACHEFINANCE:${symbol}=${attribute}. Google=${googleFinanceValue}`);

    if (attribute.toUpperCase() === "TEST") {
        return cacheFinanceTest();
    }

    if (symbol === '' || attribute === '') {
        return '';
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

        if (data !== null)
            return data;

        //  Last resort... try other sites.
        let stockAttributes = ThirdPartyFinance.get(symbol, attribute);

        //  Failed third party lookup, try using long term cache.
        if (!stockAttributes.isAttributeSet(attribute)) {
            const cachedStockAttribute = CacheFinance.getFinanceValueFromCache(cacheKey, false);
            if (cachedStockAttribute !== null)
                stockAttributes = cachedStockAttribute;
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
     * @returns {any}
     */
    static getFinanceValueFromCache(cacheKey, useShortCacheOnly) {
        const shortCache = CacheService.getScriptCache();
        const longCache = new ScriptSettings();

        let data = shortCache.get(cacheKey);

        //  Set to null while testing.  Remove when all is working.
        // data = null;

        if (data !== null && data !== "#ERROR!") {
            Logger.log(`Found in Short CACHE: ${cacheKey}. Value=${data}`);
            const parsedData = JSON.parse(data);
            if (!(typeof parsedData === 'string' && parsedData === "#ERROR!"))
                return parsedData;
        }

        if (useShortCacheOnly)
            return null;

        data = longCache.get(cacheKey);
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
        if (stockAttributes.stockName !== null)
            CacheFinance.saveFinanceValueToCache(CacheFinance.makeCacheKey(symbol, "NAME"), stockAttributes.stockName, 1200);
        if (stockAttributes.stockPrice !== null)
            CacheFinance.saveFinanceValueToCache(CacheFinance.makeCacheKey(symbol, "PRICE"), stockAttributes.stockPrice, 1200);
        if (stockAttributes.yieldPct !== null)
            CacheFinance.saveFinanceValueToCache(CacheFinance.makeCacheKey(symbol, "YIELDPCT"), stockAttributes.yieldPct, 1200);
    }

    /**
     * 
     * @param {String} key 
     * @param {any} financialData 
     * @param {Number} shortCacheSeconds 
     * @returns {void}
     */
    static saveFinanceValueToCache(key, financialData, shortCacheSeconds = 1200) {
        const shortCache = CacheService.getScriptCache();
        const longCache = new ScriptSettings();
        let start = new Date().getTime();

        const currentShortCacheValue = shortCache.get(key);
        if (currentShortCacheValue !== null && JSON.parse(currentShortCacheValue) === financialData) {
            Logger.log(`GoogleFinance VALUE.  No Change in SHORT Cache. ms=${new Date().getTime() - start}`);
            return;
        }

        if (currentShortCacheValue !== null)
            Logger.log(`Short Cache Changed.  Old=${JSON.parse(currentShortCacheValue)} . New=${financialData}`);
   
        //  If we normally get the price from Google, we want to cache for a longer
        //  time because the only time we need a price for this particular stock
        //  is when GOOGLEFINANCE fails.
        start = new Date().getTime();
        shortCache.put(key, JSON.stringify(financialData), shortCacheSeconds);
        const shortMs = new Date().getTime() - start;
       
        //  For emergency cases when GOOGLEFINANCE is down long term...
        start = new Date().getTime();
        longCache.put(key, financialData, 7);
        const longMs = new Date().getTime() - start;

        Logger.log(`SET GoogleFinance VALUE Long/Short Cache. Key=${key}.  Value=${financialData}. Short ms=${shortMs}. Long ms=${longMs}`);
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
                }

                if (propertyValue !== null && (PropertyData.isExpired(propertyValue) || deleteAll)) {
                    this.scriptProperties.deleteProperty(key);
                    Logger.log(`Removing expired SCRIPT PROPERTY: key=${key}`);
                }
            }
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
class ThirdPartyFinance {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    static get(symbol, attribute) {
        let data = new StockAttributes();

        switch (attribute) {
            case "PRICE":
                data = ThirdPartyFinance.getStockPrice(symbol);
                break;

            case "YIELDPCT":
                data = ThirdPartyFinance.getStockDividendYield(symbol);
                break;

            case "NAME":
                data = ThirdPartyFinance.getName(symbol);
                break;

            default:
                Logger.log(`3'rd Party FINANCE attribute not supported: ${attribute}`);
                break;
        }

        if (data.stockPrice !== null)
            data.stockPrice = Math.round(data.stockPrice * 100) / 100;

        if (data.yieldPct !== null)
            data.yieldPct = Math.round(data.yieldPct * 10000) / 10000;

        return data;
    }

    /**
     * 
     * @param {string} symbol 
     * @returns {StockAttributes}
     */
    static getStockPrice(symbol) {
        /** @type {StockAttributes} */
        let data = GlobeAndMail.getInfo(symbol);

        if (data.stockPrice === null)
            data = TdMarketResearch.getInfo(symbol, "ETF");

        if (data.stockPrice === null)
            data = TdMarketResearch.getInfo(symbol, "STOCK");

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getStockDividendYield(symbol) {

        let data = GlobeAndMail.getInfo(symbol);

        if (data.yieldPct === null)
            data = TdMarketResearch.getInfo(symbol, "ETF");

        if (data.yieldPct === null)
            data = TdMarketResearch.getInfo(symbol, "STOCK");

        if (data.yieldPct === null)
            data = YahooFinance.getInfo(symbol);

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getName(symbol) {
        /** @type {StockAttributes} */
        let data = GlobeAndMail.getInfo(symbol);

        if (data.stockName === null)
            data = TdMarketResearch.getInfo(symbol, "ETF");

        if (data.stockName === null)
            data = TdMarketResearch.getInfo(symbol, "STOCK");

        return data;
    }

    static getTickerCountryCode(symbol) {
        const colon = symbol.indexOf(":");
        let exchange = "";
        let countryCode = "ca";

        if (colon < 0) {
            return countryCode;
        }

        const parts = symbol.split(":");
        exchange = parts[0].toUpperCase();

        switch(exchange) {
            case "NASDAQ":
            case "NYSEARCA":
                countryCode = "us";
                break;   
        }

        return countryCode;
    }
}

class TdMarketResearch {
    /**
     * 
     * @param {String} symbol 
     * @param {String} type 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, type = "ETF") {
        const data = new StockAttributes();

        let URL = null;
        if (type === "ETF")
            URL = `https://marketsandresearch.td.com/tdwca/Public/ETFsProfile/Summary/${ThirdPartyFinance.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;
        else
            URL = `https://marketsandresearch.td.com/tdwca/Public/Stocks/Overview/${ThirdPartyFinance.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return data;
        }
        Logger.log(`getStockDividendYield:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        //  Get the dividend yield.
        let parts = html.match(/.Dividend Yield<\/th><td class="last">(\d*\.?\d*)\%/);
        if (parts === null) {
            parts = html.match(/.Dividend Yield<\/div>.*?cell-container contains">(\d*\.?\d*)%/);
        }
        if (parts !== null && parts.length === 2) {
            const tempPct = parts[1];

            const parsedValue = parseFloat(tempPct) / 100;

            if (!isNaN(parsedValue)) {
                data.yieldPct = parsedValue;
            }
        }

        //  Get the name.
        parts = html.match(/.\<span class=\"issueName">(.*?)<\//);
        if (parts !== null && parts.length === 2) {
            data.stockName = parts[1];
        }

        //  Get the price.
        parts = html.match(/.LAST PRICE<\/span<div><span>(\d*\.?\d*)</);
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
            symbol = `${symbol.substr(0, dash)}.PR.${symbol.substr(dash + 1)}`;
        }

        return symbol;
    }
}

class YahooFinance {
    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        const data = new StockAttributes();

        const URL = `https://finance.yahoo.com/quote/${YahooFinance.getTicker(symbol)}`;

        const html = UrlFetchApp.fetch(URL).getContentText();
        Logger.log(`getStockDividendYield:  ${symbol}`);
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

class GlobeAndMail {
    /**
     * Only gets dividend yield.
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        const data = new StockAttributes();
        const URL = `https://www.theglobeandmail.com/investing/markets/stocks/${GlobeAndMail.getTicker(symbol)}`;

        Logger.log(`getStockDividendYield:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return data;
        }

        //  Get the dividend yield.
        let parts = html.match(/.name=\"dividendYieldTrailing".*?value="(\d*\.?\d*)%/);

        if (parts === null)
            parts = html.match(/.name=\\"dividendYieldTrailing\\".*?value=\\"(\d*\.?\d*)%/);

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
        parts = html.match(/."lastPrice":"(\d*\.?\d*)"/);
        if (parts !== null && parts.length === 2) {

            const parsedValue = parseFloat(parts[1]);
            if (!isNaN(parsedValue)) {
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
 * Get/Set data about stocks/ETFs.
 */
class StockAttributes {
    constructor() {
        this.yieldPct = null;
        this.stockName = null;
        this.stockPrice = null;
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
 * Returns a diagnostic list of 3rd party stock lookup info.
 * @returns {any[][]}
 */
function cacheFinanceTest() {
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
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "badSymbol");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "ZTL");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "TSE:MEG");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "NASDAQ:MSFT");
        
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "NYSEARCA:SHYG");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "ZTL");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:ZTL");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:DFN-A");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "ZTL", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:DFN-A", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "NASDAQ:MSFT", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "badSymbol");

        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NYSEARCA:SHYG");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:FTN-A");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:HBF.B");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:MEG");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:ZTL");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NASDAQ:MSFT");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "badSymbol");

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
     */
    run(serviceName, func, symbol, type = "ETF") {
        const result = new CacheFinanceTestStatus(serviceName, symbol);
        try {
            /** @type {StockAttributes} */
            const data = func(symbol, type);
            result.setStatus("ok");
            result.setStockAttributes(data);
            result.setTypeLookup(type);

            if (data.stockName === null && data.stockPrice === null && data.yieldPct === null) {
                result.setStatus("Not Found!")
            }
        }
        catch(ex) {
            result.setStatus("Error");
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

        let row = ["Service", "Symbol", "Status", "Price", "Yield", "Name", "Type", "Run Time(ms)"];
        resultTable.push(row);

        for (let testRun of this.testRuns) {
            let row = [];

            row.push(testRun.serviceName);
            row.push(testRun.symbol);
            row.push(testRun.status);
            row.push(testRun.stockAttributes.stockPrice);
            row.push(testRun.stockAttributes.yieldPct);
            row.push(testRun.stockAttributes.stockName);
            row.push(testRun.typeLookup);
            row.push(testRun.runTime);

            resultTable.push(row);
        }

        return resultTable;
    }
}

/**
 * Individual test results and tracking.
 */
class CacheFinanceTestStatus {
    constructor (serviceName="", symbol="") {
        this._serviceName = serviceName;
        this._symbol = symbol;
        this._stockAttributes = new StockAttributes();
        this._startTime = Date.now()
        this._typeLookup = "";
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

    setServiceName(val) {
        this._serviceName = val;
        return this;
    }
    setSymbol(val) {
        this._symbol = val;
        return this;
    }
    setValue(val) {
        this._value = val;
        return this;
    }
    setStatus(val) {
        this._status = val;
        return this;
    }
    setRunTime(val) {
        this._runTime = val;
        return this;
    }
    setTypeLookup(val) {
        this._typeLookup = val;
        return this;
    }
    /**
     * 
     * @param {StockAttributes} val 
     */
    setStockAttributes(val) {
        this._stockAttributes = val;
    }
    finishTimer() {
        this._runTime = Date.now() - this._startTime;
    }
}



