/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { TdMarketResearch, GlobeAndMail, YahooFinance, StockAttributes, FinnHub } from "./CacheFinanceWebSites.js";
import { ThirdPartyFinance, FinanceWebsiteSearch } from "./CacheFinance3rdParty.js";
import { CACHEFINANCE, CacheFinance } from "./CacheFinance.js";
export { cacheFinanceTest };

class Logger {
    static log(msg) {
        console.log(msg);
    }
}
//  *** DEBUG END ***/

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
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "ZTL", "ALL", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:DFN-A", "ALL", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "NASDAQ:MSFT", "ALL", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:RY", "ALL", "STOCK");

        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NYSEARCA:SHYG");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:FTN-A");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:HBF.B");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:MEG");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:RY");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NASDAQ:MSFT");

        CacheFinance.deleteFromCache("TSE:RY", "PRICE");
        this.cacheTestRun.run("CACHEFINANCE - not cached", CACHEFINANCE, "TSE:RY", "PRICE", "##NotSet##");
        this.cacheTestRun.run("CACHEFINANCE - cached", CACHEFINANCE, "TSE:RY", "PRICE", "##NotSet##");

        const plan = new FinanceWebsiteSearch();
        //  Make fresh lookup plans.
        plan.deleteLookupPlan("TSE:RY");
        plan.deleteLookupPlan("NASDAQ:BNDX");
        plan.deleteLookupPlan("TSE:ZTL");

        plan.getLookupPlan("TSE:RY", "");
        plan.getLookupPlan("NASDAQ:BNDX", "");
        plan.getLookupPlan("TSE:ZTL", "");

        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "TSE:RY", "PRICE");
        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "NASDAQ:BNDX", "PRICE");
        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "NASDAQ:BNDX", "NAME");
        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "NASDAQ:BNDX", "YIELDPCT");
        this.cacheTestRun.run("OptimalSite", ThirdPartyFinance.get, "TSE:ZTL", "PRICE");

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
            result.setStatus("Error: " + ex);
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

