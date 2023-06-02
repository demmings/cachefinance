/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { ScriptSettings } from "./SQL/ScriptSettings.js";
import { CacheFinanceWebSites, StockAttributes, FinanceWebSite } from "./CacheFinanceWebSites.js";
export { ThirdPartyFinance };

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

        if (data.stockPrice !== null)
            data.stockPrice = Math.round(data.stockPrice * 100) / 100;

        if (data.yieldPct !== null)
            data.yieldPct = Math.round(data.yieldPct * 10000) / 10000;

        return data;
    }
}

/**
 * @classdesc Make a plan to find finance data from websites and then execute the plan and return the data.
 */
class FinanceWebsiteSearch {
    constructor() {
        const siteInfo = new CacheFinanceWebSites();

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

}

/**
 * @classdesc Ordered list of sites for doing attribute lookup.
 * This object will be converted to JSON for storage.
 */
class FinanceSiteList {
    constructor(symbol) {
        this.symbol = symbol;
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

        this.priceSites = this.priceSites.sort((a,b) => a.timeMs - b.timeMs);
        this.nameSites = this.nameSites.sort((a,b) => a.timeMs - b.timeMs);
        this.yieldSites = this.yieldSites.sort((a,b) => a.timeMs - b.timeMs);
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
     * @returns 
     */
    static getStockAttribute(stockSites, attribute) {
        let data = new StockAttributes();
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

        const sitesSearchFunction = new CacheFinanceWebSites();
        for (const site of siteArr) {
            const siteFunction = sitesSearchFunction.getByName(site);
            data = siteFunction.getInfo(stockSites.symbol);

            if (data !== null) {
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