import { describe, expect, it } from "vitest";
import { ScriptSettings, PropertyData } from "../src/ScriptSettings.js";

describe("PropertyData", () => {
    it("stores and retrieves JSON-serializable values", () => {
        const property = new PropertyData({ price: 12.34 }, 1);
        expect(PropertyData.getData(property)).toEqual({ price: 12.34 });
    });

    it("reports expiry based on days to hold", () => {
        const property = new PropertyData("stale", 0);
        property.expiry = Date.now() - 1000;
        expect(PropertyData.isExpired(property)).toBe(true);
    });
});

describe("ScriptSettings", () => {
    it("stores and retrieves values before expiry", () => {
        const settings = new ScriptSettings();
        settings.put("PRICE|TSE:ZTL", 10.5, 1);

        expect(settings.get("PRICE|TSE:ZTL")).toBe(10.5);
    });

    it("returns null for missing keys", () => {
        const settings = new ScriptSettings();
        expect(settings.get("MISSING|KEY")).toBeNull();
    });

    it("retrieves multiple keys with one properties lookup", () => {
        ScriptSettings.putAllKeysWithData(
            ["PRICE|TSE:A", "PRICE|TSE:B"],
            [10, 20],
            1
        );

        expect(ScriptSettings.getAll(["PRICE|TSE:A", "PRICE|TSE:B", "PRICE|TSE:C"]))
            .toEqual([10, 20, null]);
    });

    it("deletes a specific key", () => {
        const settings = new ScriptSettings();
        settings.put("PRICE|TSE:ZTL", 10.5, 1);
        settings.delete("PRICE|TSE:ZTL");

        expect(settings.get("PRICE|TSE:ZTL")).toBeNull();
    });

    it("returns null and deletes expired values on read", () => {
        const settings = new ScriptSettings();
        settings.put("PRICE|TSE:ZTL", 10.5, 1);

        const raw = settings.scriptProperties.getProperty("PRICE|TSE:ZTL");
        const parsed = JSON.parse(raw);
        parsed.expiry = Date.now() - 1000;
        settings.scriptProperties.setProperty("PRICE|TSE:ZTL", JSON.stringify(parsed));

        expect(settings.get("PRICE|TSE:ZTL")).toBeNull();
        expect(settings.scriptProperties.getProperty("PRICE|TSE:ZTL")).toBeNull();
    });

    it("removes expired properties during expire cleanup", () => {
        const settings = new ScriptSettings();
        settings.put("PRICE|TSE:OLD", 1, 1);

        const raw = settings.scriptProperties.getProperty("PRICE|TSE:OLD");
        const parsed = JSON.parse(raw);
        parsed.expiry = Date.now() - 1000;
        settings.scriptProperties.setProperty("PRICE|TSE:OLD", JSON.stringify(parsed));

        ScriptSettings.expire(false, 10);

        expect(settings.get("PRICE|TSE:OLD")).toBeNull();
    });

    it("removes all cache properties when deleteAll is true", () => {
        const settings = new ScriptSettings();
        settings.put("PRICE|TSE:A", 10, 30);
        settings.put("PRICE|TSE:B", 20, 30);

        ScriptSettings.expire(true);

        expect(settings.get("PRICE|TSE:A")).toBeNull();
        expect(settings.get("PRICE|TSE:B")).toBeNull();
    });

    it("handles invalid JSON in script properties during expire", () => {
        const settings = new ScriptSettings();
        settings.scriptProperties.setProperty("NOT_CACHEFINANCE", "plain-text");

        expect(() => ScriptSettings.expire(false, 1)).not.toThrow();
        expect(settings.scriptProperties.getProperty("NOT_CACHEFINANCE")).toBe("plain-text");
    });
});
