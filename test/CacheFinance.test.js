import { beforeEach, describe, expect, it, vi } from "vitest";
import { CACHEFINANCE, CACHEFINANCES, CacheFinance } from "../src/CacheFinance.js";
import { ThirdPartyFinance } from "../src/CacheFinance3rdParty.js";
import { StockAttributes } from "../src/CacheFinanceWebSites.js";
import { CacheFinanceUtils } from "../src/CacheFinanceUtils.js";
import { CacheService } from "../src/GasMocks.js";

function mockThirdPartyLookup(valuesByAttribute = {}) {
    vi.spyOn(ThirdPartyFinance, "getMissingStockAttributesFromThirdParty")
        .mockImplementation((symbols, attribute) => symbols.map(() => {
            const data = new StockAttributes();
            const value = valuesByAttribute[attribute];

            if (attribute === "PRICE") {
                data.stockPrice = value ?? 42.5;
            }
            else if (attribute === "NAME") {
                data.stockName = value ?? "Test Corp";
            }
            else if (attribute === "YIELDPCT") {
                data.yieldPct = value ?? 0.05;
            }

            return data;
        }));
}

describe("CacheFinance helpers", () => {
    beforeEach(() => {
        vi.restoreAllMocks();
    });

    it("identifies symbols missing valid default values", () => {
        expect(CacheFinance.getSymbolsWithNoValidData(
            ["TSE:A", "TSE:B", "TSE:C"],
            [10, "#N/A", ""]
        )).toEqual(["TSE:B", "TSE:C"]);
    });

    it("deduplicates symbols with missing data", () => {
        expect(CacheFinance.getSymbolsWithNoValidData(
            ["TSE:ZTL", "TSE:FTN-A", "TSE:ZTL"],
            ["#N/A", 10, null]
        )).toEqual(["TSE:ZTL"]);
    });

    it("merges third-party values into duplicate symbol positions", () => {
        const merged = CacheFinance.updateMasterWithMissed(
            ["TSE:ZTL", "TSE:FTN-A", "TSE:ZTL"],
            [null, 10, null],
            ["TSE:ZTL"],
            [15]
        );

        expect(merged).toEqual([15, 10, 15]);
    });

    it("maps stock attributes back to finance values", () => {
        const attrs = [new StockAttributes(), new StockAttributes()];
        attrs[0].stockPrice = 12.34;
        attrs[1].stockName = "Example ETF";

        expect(CacheFinance.getValuesFromStockAttributes(attrs, "PRICE")).toEqual([12.34, ""]);
        expect(CacheFinance.getValuesFromStockAttributes(attrs, "NAME")).toEqual(["", "Example ETF"]);
    });

    it("reads a value from the short cache", () => {
        const key = "PRICE|TSE:ZTL";
        CacheService.getScriptCache().put(key, JSON.stringify(88.8), 1200);

        expect(CacheFinance.getFinanceValueFromShortCache(key)).toBe(88.8);
        expect(CacheFinance.getFinanceValueFromShortCache("MISSING|KEY")).toBeNull();
    });

    it("ignores short-cache entries that contain errors", () => {
        const key = "PRICE|TSE:ZTL";
        CacheService.getScriptCache().put(key, "#ERROR!", 1200);
        expect(CacheFinance.getFinanceValueFromShortCache(key)).toBeNull();

        CacheService.getScriptCache().put(key, JSON.stringify("#ERROR!"), 1200);
        expect(CacheFinance.getFinanceValueFromShortCache(key)).toBeNull();
    });

    it("skips short-cache reads when all google values are valid", () => {
        const spy = vi.spyOn(CacheFinanceUtils, "bulkShortCacheGet");

        const values = CacheFinance.updateMissingValuesFromShortCache(
            ["TSE:A"],
            "PRICE",
            [12.5]
        );

        expect(values).toEqual([12.5]);
        expect(spy).not.toHaveBeenCalled();
    });

    it("reports when all default values are valid", () => {
        expect(CacheFinance.isAllGoogleDefaultValuesValid(["TSE:A"], [10])).toBe(true);
        expect(CacheFinance.isAllGoogleDefaultValuesValid(["TSE:A"], ["#N/A"])).toBe(false);
        expect(CacheFinance.isAllGoogleDefaultValuesValid(["TSE:A", "TSE:B"], [10])).toBe(false);
    });

    it("deletes short and long cache entries for a symbol", () => {
        const settingsKey = "PRICE|TSE:ZTL";
        CacheFinanceUtils.putFinanceValuesIntoShortCache([settingsKey], [10], 1200);
        CacheFinanceUtils.bulkLongCachePut(["TSE:ZTL"], "PRICE", [10], 1);

        CacheFinance.deleteFromCache("TSE:ZTL", "PRICE");

        expect(CacheFinance.getFinanceValueFromShortCache(settingsKey)).toBeNull();
        expect(CacheFinanceUtils.bulkLongCacheGet(["TSE:ZTL"], "PRICE")).toEqual([null]);
    });
});

describe("CacheFinance.getBulkFinanceData", () => {
    beforeEach(() => {
        vi.restoreAllMocks();
    });

    it("returns valid google finance values as a column range", () => {
        const result = CacheFinance.getBulkFinanceData(
            ["TSE:A", "TSE:B"],
            "PRICE",
            [10.5, 20.25]
        );

        expect(result).toEqual([[10.5], [20.25]]);
    });

    it("fills missing values from the short cache", () => {
        CacheFinanceUtils.putFinanceValuesIntoShortCache(
            ["PRICE|TSE:ZTL"],
            [55.5],
            1200
        );

        const result = CacheFinance.getBulkFinanceData(
            ["TSE:ZTL"],
            "PRICE",
            ["#N/A"]
        );

        expect(result).toEqual([[55.5]]);
    });

    it("uses third-party lookups when cache and defaults are missing", () => {
        mockThirdPartyLookup({ PRICE: 77.7 });

        const result = CacheFinance.getBulkFinanceData(
            ["TSE:ZTL"],
            "PRICE",
            ["#N/A"]
        );

        expect(result).toEqual([[77.7]]);
    });

    it("falls back to long cache values as a last resort", () => {
        vi.spyOn(ThirdPartyFinance, "getMissingStockAttributesFromThirdParty")
            .mockImplementation((symbols) => symbols.map(() => new StockAttributes()));
        CacheFinanceUtils.bulkLongCachePut(["TSE:ZTL"], "PRICE", [33.3], 1);

        const result = CacheFinance.getBulkFinanceData(
            ["TSE:ZTL"],
            "PRICE",
            ["#N/A"]
        );

        expect(result).toEqual([[33.3]]);
    });
});

describe("CacheFinance backdoor commands", () => {
    beforeEach(() => {
        vi.restoreAllMocks();
    });

    it("returns help text for ?", () => {
        const help = CacheFinance.backDoorCommands("", "", "?", "");
        expect(help[0][0]).toContain("Valid commands");
    });

    it("lists supported providers", () => {
        const providers = CacheFinance.backDoorCommands("", "", "LIST", "");
        expect(providers.flat()).toEqual(expect.arrayContaining(["YAHOOAPI", "FINNHUB", "GLOBE"]));
    });

    it("stores and reads preferred providers", () => {
        CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "SET", "YAHOO");
        expect(CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "GET", ""))
            .toBe("YAHOO");
    });

    it("rejects invalid provider names", () => {
        expect(CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "SET", "NOT_A_SITE"))
            .toBe("Invalid provider name.  No change made.");
    });

    it("clears cache for a specific symbol and attribute", () => {
        CacheFinanceUtils.putFinanceValuesIntoShortCache(["PRICE|TSE:CJP"], [12], 1200);
        expect(CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "CLEARCACHE", ""))
            .toBe("Cache Cleared");
        expect(CacheFinance.getFinanceValueFromShortCache("PRICE|TSE:CJP")).toBeNull();
    });

    it("clears all cache entries when symbol and attribute are blank", () => {
        CacheFinanceUtils.bulkLongCachePut(["TSE:A"], "PRICE", [10], 1);
        expect(CacheFinance.backDoorCommands("", "", "CLEARCACHE", "")).toBe("Cache Cleared");
        expect(CacheFinanceUtils.bulkLongCacheGet(["TSE:A"], "PRICE")).toEqual([null]);
    });

    it("removes old cache entries with EXPIRECACHE", () => {
        CacheFinanceUtils.bulkLongCachePut(["TSE:OLD"], "PRICE", [5], 1);
        expect(CacheFinance.backDoorCommands("", "", "EXPIRECACHE", ""))
            .toBe("Old Cache Items Removed");
    });

    it("moves the preferred provider to the blocked list on REMOVE", () => {
        CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "SET", "YAHOO");
        const message = CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "REMOVE", "");

        expect(message).toContain("Site removed");
        expect(CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "GET", ""))
            .toBe("No site set.");
        expect(CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "GETBLOCKED", ""))
            .toBe("YAHOO");
    });

    it("stores and reads blocked providers", () => {
        CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "SETBLOCKED", "GLOBE");
        expect(CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "GETBLOCKED", ""))
            .toBe("GLOBE");
    });

    it("clears a provider preference when SET is called with an empty site", () => {
        CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "SET", "YAHOO");
        CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "SET", "");
        expect(CacheFinance.backDoorCommands("TSE:CJP", "PRICE", "GET", ""))
            .toBe("No site set.");
    });

    it("reports when no preferred provider exists during REMOVE", () => {
        expect(CacheFinance.backDoorCommands("TSE:NEW", "PRICE", "REMOVE", ""))
            .toBe("Currently no preferred site for TSE:NEW PRICE");
    });
});

describe("CACHEFINANCE custom function", () => {
    beforeEach(() => {
        vi.restoreAllMocks();
        mockThirdPartyLookup({ PRICE: 99.1 });
    });

    it("returns a single looked-up value", () => {
        expect(CACHEFINANCE("TSE:ZTL", "price", "#N/A")).toBe(99.1);
    });

    it("returns an empty string for blank symbol or attribute", () => {
        expect(CACHEFINANCE("", "price")).toBe("");
        expect(CACHEFINANCE("TSE:ZTL", "")).toBe("");
    });
});

describe("CACHEFINANCES custom function", () => {
    beforeEach(() => {
        vi.restoreAllMocks();
    });

    it("returns a 2D array for range lookups", () => {
        const result = CACHEFINANCES([["TSE:A"], ["TSE:B"]], "price", [[10], [20]]);
        expect(result).toEqual([[10], [20]]);
    });

    it("returns a scalar for single-value lookups", () => {
        expect(CACHEFINANCES("TSE:A", "price", 10)).toBe(10);
    });

    it("throws when symbol and default ranges differ in size", () => {
        expect(() => CACHEFINANCES([["TSE:A"], ["TSE:B"]], "price", [[10]]))
            .toThrow("Stock symbol RANGE must match default values range.");
    });

    it("returns an empty string when no symbols are provided", () => {
        expect(CACHEFINANCES([], "price", [])).toBe("");
    });
});
