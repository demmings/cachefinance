import { beforeEach, describe, expect, it } from "vitest";
import { CacheService } from "../src/GasMocks.js";
import { SiteThrottle, ThresholdPeriod } from "../src/CacheFinanceUtils.js";
import { ScriptSettings } from "../src/ScriptSettings.js";

describe("SiteThrottle", () => {
    beforeEach(() => {
        CacheService.reset();
    });

    it("allows requests below the per-second limit", () => {
        const throttle = new SiteThrottle("TESTSITE", [
            new ThresholdPeriod("SECOND", 3)
        ]);

        expect(throttle.checkAndIncrement()).toBe(true);
        expect(throttle.checkAndIncrement()).toBe(true);
        expect(throttle.checkAndIncrement()).toBe(false);
    });

    it("persists updated counters to the short cache", () => {
        const throttle = new SiteThrottle("TESTSITE", [
            new ThresholdPeriod("SECOND", 5)
        ]);

        throttle.checkAndIncrement();
        throttle.update();

        const key = SiteThrottle.createSecondKey("TESTSITE");
        expect(SiteThrottle.currentForSecond(key)).toBe(1);
    });

    it("tracks day-based limits in script settings", () => {
        const throttle = new SiteThrottle("DAYSITE", [
            new ThresholdPeriod("DAY", 2)
        ]);

        expect(throttle.checkAndIncrement()).toBe(true);
        throttle.update();
        expect(throttle.checkAndIncrement()).toBe(false);
    });

    it("builds stable throttle keys for each interval", () => {
        expect(SiteThrottle.makeKey("SITE", "MIN", 12)).toBe("SITE:MIN:12");
        expect(SiteThrottle.createMinuteKey("SITE")).toMatch(/^SITE:MIN:\d+$/);
        expect(SiteThrottle.createDayKey("SITE")).toMatch(/^SITE:DAY:\d+$/);
        expect(SiteThrottle.createMonthKey("SITE")).toMatch(/^SITE:MONTH:\d+$/);
    });

    it("throws for unsupported threshold periods", () => {
        expect(() => SiteThrottle.getCurrentThresholds(
            [new ThresholdPeriod("WEEK", 1)],
            "SITE"
        )).toThrow("Invalid threshold period WEEK");
    });
});

describe("ThresholdPeriod", () => {
    it("stores period name and max request count", () => {
        const period = new ThresholdPeriod("MINUTE", 8);
        period.periodName = "SECOND";
        period.maxPerPeriod = 10;

        expect(period.periodName).toBe("SECOND");
        expect(period.maxPerPeriod).toBe(10);
    });
});
