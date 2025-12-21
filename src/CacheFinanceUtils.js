/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { ScriptSettings } from "./SQL/ScriptSettings.js";
export { CacheFinanceUtils, SiteThrottle, ThresholdPeriod };

class Logger {
    static log(msg) {
        console.log(msg);
    }
}
//  *** DEBUG END ***/

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

                case "MONTH":
                    key = SiteThrottle.createMonthKey(siteID);
                    current = SiteThrottle.currentForMonth(key);
                    Logger.log(`MONTH Check.  key=${key}. Current=${current.toString()}.`);
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

                case "MONTH":
                    SiteThrottle.updateForMonth(this.periodKeys[i], this.periodCount[i]);
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
     * @param {String} key 
     * @param {Number} current 
     */
    static updateForMonth(key, current) {
        const longCache = new ScriptSettings();
        longCache.put(key, current, 30);
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
     * Current requests made for the day.
     * @param {String} key 
     * @returns {Number}
     */
    static currentForMonth(key) {
        //  For now it is the same implentation as the DAY, but I want a separate function in case they differ in future.
        return SiteThrottle.currentForDay(key);
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

    /**
     * 
     * @param {String} siteID 
     * @returns {String}
     */
    static createMonthKey(siteID) {
        //  Month throttle will only really work if around the same number of request happen per day.
        //  On the first day of a new month, the count will be zero - which does not take into account
        //  the last 29 days of the previous month.  More work is needed for a rolling throttle....
        const today = new Date();
        const dayNum = today.getMonth();      // get month #
        return SiteThrottle.makeKey(siteID, "MONTH", dayNum);
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