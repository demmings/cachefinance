import { beforeEach } from "vitest";
import { CacheService, PropertiesService, SpreadsheetApp, UrlFetchApp } from "../src/GasMocks.js";

globalThis.CacheService = CacheService;
globalThis.PropertiesService = PropertiesService;
globalThis.UrlFetchApp = UrlFetchApp;
globalThis.SpreadsheetApp = SpreadsheetApp;
globalThis.Logger = {
    log() {}
};
globalThis.CacheFinance = {
    deleteFromShortCache() {}
};

beforeEach(() => {
    CacheService.reset();
    PropertiesService.reset();
    UrlFetchApp.reset();
});
