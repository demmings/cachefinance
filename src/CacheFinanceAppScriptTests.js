/**
 * Optional Apps Script helpers for manual authorization and debugging.
 * Not included in dist/CacheFinance.js — paste into a separate script file in your
 * Google Sheets project if you want Run-menu test entry points.
 */

// skipcq: JS-0128
function testYieldPct() {
    const val = CACHEFINANCE("TSE:CJP", "yieldpct");        // skipcq: JS-0128
    Logger.log(`Test CacheFinance TSE:CJP(yieldpct)=${val}`);
}

function testCacheFinances() {                                  // skipcq: JS-0128
    const symbols = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("A30:A165").getValues();
    const data = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("E30:E165").getValues();

    const cacheData = CACHEFINANCES(symbols, "PRICE", data);

    const singleSymbols = CacheFinanceUtils.convertRowsToSingleArray(symbols);

    Logger.log(`BULK CACHE TEST Success${cacheData} . ${singleSymbols}`);
}

function testUpdateMaster() {
    const symbols = ["TSE:ZTL", "TSE:FTN-A", "TSE:ZTL"];
    const googleFinanceValues = [null, 10.0, null];
    const symbolsWithNoData = ["TSE:ZTL"];
    const thirdPartyFinanceValues = [15];

    const newGoogleFinance = CacheFinance.updateMasterWithMissed(symbols, googleFinanceValues, symbolsWithNoData, thirdPartyFinanceValues);
    Logger.log(newGoogleFinance);
}
