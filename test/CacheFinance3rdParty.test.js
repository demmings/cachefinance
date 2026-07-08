import { beforeEach, describe, expect, it } from "vitest";
import {
    FinanceSiteLookupAnalyzer,
    FinanceSiteLookupStats,
    FinanceWebsiteSearch,
    StockWebURL
} from "../src/CacheFinance3rdParty.js";
import { FinanceWebSite, StockAttributes } from "../src/CacheFinanceWebSites.js";
import { ScriptSettings } from "../src/ScriptSettings.js";
import { UrlFetchApp } from "../src/GasMocks.js";

function buildParseResponse(price) {
    return (_html, _symbol, attribute) => {
        const data = new StockAttributes();
        if (attribute === "PRICE") {
            data.stockPrice = price;
        }
        return data;
    };
}

describe("StockWebURL", () => {
    it("prioritizes the preferred provider URL", () => {
        const stock = new StockWebURL("TSE:ZTL");
        stock.addSiteURL("GLOBE", "", "", "http://globe", buildParseResponse(1), null);
        stock.addSiteURL("YAHOO", "YAHOO", "", "http://yahoo", buildParseResponse(2), null);

        expect(stock.getURL()).toBe("http://yahoo");
        expect(stock.siteName[0]).toBe("YAHOO");
    });

    it("skips blocked sites and empty URLs", () => {
        const stock = new StockWebURL("TSE:ZTL");
        stock.addSiteURL("YAHOO", "", "YAHOO", "http://yahoo", buildParseResponse(2), null);
        stock.addSiteURL("GLOBE", "", "", "http://globe", buildParseResponse(1), null);

        expect(stock.getURL()).toBe("http://globe");
    });

    it("parses responses and advances to the next site", () => {
        const stock = new StockWebURL("TSE:ZTL");
        stock.addSiteURL("YAHOO", "", "", "http://yahoo", buildParseResponse(12.34), null);

        const data = stock.parseResponse("<html>", "PRICE");
        expect(data.stockPrice).toBe(12.34);
        expect(data.isAttributeSet("PRICE")).toBe(true);

        stock.skipToNextSite();
        expect(stock.isSitesDone()).toBe(true);
    });

    it("records the working provider for future lookups", () => {
        const stock = new StockWebURL("TSE:ZTL");
        stock.addSiteURL("YAHOO", "", "", "http://yahoo", buildParseResponse(12.34), null);
        stock.parseResponse("<html>", "PRICE");

        const bestStockSites = {};
        stock.updateBestSites(bestStockSites, "PRICE");

        expect(bestStockSites["PRICE|TSE:ZTL"]).toBe("YAHOO");
    });
});

describe("FinanceWebsiteSearch", () => {
    beforeEach(() => {
        UrlFetchApp.reset();
    });

    it("creates cache keys for lookup plans", () => {
        expect(FinanceWebsiteSearch.makeCacheKey("TSE:ZTL")).toBe("WebSearch|TSE:ZTL");
    });

    it("reads and writes preferred provider metadata", () => {
        FinanceWebsiteSearch.writeBestStockWebsites({ "PRICE|TSE:ZTL": "YAHOO" });
        expect(FinanceWebsiteSearch.readBestStockWebsites()).toEqual({ "PRICE|TSE:ZTL": "YAHOO" });
    });

    it("builds the next URL batch from pending stock lookups", () => {
        const stock = new StockWebURL("TSE:ZTL");
        stock.addSiteURL("YAHOO", "", "", "http://yahoo", buildParseResponse(10), null);
        stock.addSiteURL("GLOBE", "", "", "http://globe", buildParseResponse(11), null);

        const [urls, batch] = FinanceWebsiteSearch.getNextUrlBatch([stock]);

        expect(urls).toEqual(["http://yahoo"]);
        expect(batch).toHaveLength(1);
    });

    it("fetches and applies website responses in bulk", () => {
        const stock = new StockWebURL("TSE:ZTL");
        stock.addSiteURL("YAHOO", "", "", "http://yahoo", buildParseResponse(15.5), null);
        const bestStockSites = {};

        UrlFetchApp.mockResponse("http://yahoo", "<html>");

        FinanceWebsiteSearch.updateStockResults(
            [stock],
            ["http://yahoo"],
            ["<html>"],
            "PRICE",
            bestStockSites
        );

        expect(stock.stockAttributes.stockPrice).toBe(15.5);
        expect(bestStockSites["PRICE|TSE:ZTL"]).toBe("YAHOO");
    });

    it("returns empty results when no symbols are requested", () => {
        expect(FinanceWebsiteSearch.getAll([], "PRICE")).toEqual([]);
    });

    it("fetches multiple URLs through UrlFetchApp.fetchAll", () => {
        UrlFetchApp.mockResponse("http://one", "one");
        UrlFetchApp.mockResponse("http://two", "two");

        expect(FinanceWebsiteSearch.bulkSiteFetch(["http://one", "http://two", ""]))
            .toEqual(["one", "two"]);
    });

    it("builds stock lookup objects for each symbol", () => {
        const bestStockSites = { "PRICE|TSE:ZTL": "YAHOO" };
        const stockUrls = FinanceWebsiteSearch.getAllStockWebSiteFunctions(
            ["TSE:ZTL"],
            "PRICE",
            bestStockSites
        );

        expect(stockUrls).toHaveLength(1);
        expect(stockUrls[0].symbol).toBe("TSE:ZTL");
        expect(stockUrls[0].siteName[0]).toBe("YAHOO");
    });

    it("skips throttled providers when building the next URL batch", () => {
        const stock = new StockWebURL("TSE:ZTL");
        const throttle = {
            checkAndIncrement: () => false,
            update: () => {}
        };

        stock.addSiteURL("YAHOO", "", "", "http://yahoo", buildParseResponse(10), throttle);
        stock.addSiteURL("GLOBE", "", "", "http://globe", buildParseResponse(11), null);

        const [urls] = FinanceWebsiteSearch.getNextUrlBatch([stock]);
        expect(urls).toEqual(["http://globe"]);
    });

    it("deletes cached lookup plans", () => {
        const settings = new ScriptSettings();
        settings.put("WebSearch|TSE:ZTL", { priceSites: ["YAHOO"] }, 1);

        FinanceWebsiteSearch.deleteLookupPlan("TSE:ZTL");

        expect(settings.get("WebSearch|TSE:ZTL")).toBeNull();
    });
});

describe("FinanceSiteLookupAnalyzer", () => {
    it("orders sites by response time and extracts the fastest values", () => {
        const slowSite = new FinanceWebSite("Slow", {});
        const fastSite = new FinanceWebSite("Fast", {});

        const slowAttrs = new StockAttributes();
        slowAttrs.stockPrice = 20;
        const fastAttrs = new StockAttributes();
        fastAttrs.stockPrice = 10;

        const slowStats = new FinanceSiteLookupStats("TSE:ZTL", slowSite)
            .setSearchTime(500)
            .setAttributes(slowAttrs);
        const fastStats = new FinanceSiteLookupStats("TSE:ZTL", fastSite)
            .setSearchTime(100)
            .setAttributes(fastAttrs);

        const analyzer = new FinanceSiteLookupAnalyzer("TSE:ZTL");
        analyzer.analyzeSiteStatus([slowStats, fastStats]);

        const siteList = analyzer.createFinanceSiteList();
        expect(siteList.priceSites).toEqual(["FAST", "SLOW"]);

        const attributes = analyzer.getStockAttributes();
        expect(attributes.stockPrice).toBe(10);
    });

    it("selects attribute-specific site lists from stored plans", () => {
        const plan = {
            symbol: "TSE:ZTL",
            priceSites: ["Fast"],
            nameSites: ["NameSite"],
            yieldSites: ["YieldSite"]
        };

        const priceData = FinanceSiteLookupAnalyzer.getStockAttribute(plan, "PRICE");
        expect(priceData).toBeInstanceOf(StockAttributes);
    });
});
