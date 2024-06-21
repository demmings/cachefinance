/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { ScriptSettings } from "./SQL/ScriptSettings.js";
import { FinanceWebSites, StockAttributes, FinanceWebSite } from "./CacheFinanceWebSites.js";
import { CacheFinanceUtils } from "./CacheFinanceUtils.js";
export { ThirdPartyFinance, FinanceWebsiteSearch };

class Logger {
    static log(msg) {
        console.log(msg);
    }
}
//  *** DEBUG END ***/

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
            const URLs = FinanceWebsiteSearch.getNextUrlBatch(missingStockData);

            Logger.log(`Batch=${batch}. URLs${URLs}`);
            const responses = FinanceWebsiteSearch.bulkSiteFetch(URLs);
            const elapsedTime = Date.now() - startTime;
            Logger.log(`Batch=${batch}. Responses=${responses.length}. Total Elapsed=${elapsedTime}`);
            batch++;

            FinanceWebsiteSearch.updateStockResults(missingStockData, URLs, responses, attribute, bestStockSites);
            
            missingStockData = missingStockData.filter(stock => ! stock.stockAttributes.isAttributeSet(attribute) && ! stock.isSitesDone())
        }

        //  Note:  If separate CACHEFINANCES() run at the same time, the last process to finish will overwrite any new results
        //         from the other runs.  This is not critical, since it is ONLY used to improve the ordering of sites to call AND
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
            matchingSites.forEach(site => site.parseResponse(responses[i], attribute));
            matchingSites.forEach(site => site.updateBestSites(bestStockSites, attribute));
            matchingSites.forEach(site => site.skipToNextSite());
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
            return {
                'url': url,         // skipcq: JS-0240 
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

/**
 * @classdesc Tracks and generates URLs to find stock data.  Also tracks parsing functions for extracting data from HTML.
 */
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