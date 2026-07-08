import fs from "node:fs";
import { describe, expect, it } from "vitest";
import {
    buildCacheFinanceBundle,
    enableDebugExports,
    hashSections,
    isCacheFinanceBundleCurrent,
    paths,
    stripDebugBlocks,
    validateBundleContent,
    writeCacheFinanceBundle
} from "../scripts/gas-source.mjs";

describe("gas-source transforms", () => {
    it("removes DEBUG blocks from source files", () => {
        const source = `/*  *** DEBUG START ***
// import { Foo } from "./Foo.js";
// export { Bar };
//  *** DEBUG END ***/

class Bar {}`;

        expect(stripDebugBlocks(source)).toBe("class Bar {}");
    });

    it("uncomments DEBUG exports for Node tests", () => {
        const source = `/*  *** DEBUG START ***
//  Remove comments for testing in NODE
// export { Widget };
//  *** DEBUG END ***/

class Widget {}`;

        const transformed = enableDebugExports(source);
        expect(transformed).toContain("export { Widget };");
        expect(transformed).toContain("class Widget {}");
    });

    it("creates stable hashes for unchanged source sections", () => {
        const sections = ["class A {}", "class B {}"];
        expect(hashSections(sections)).toBe(hashSections([...sections]));
    });
});

describe("cachefinance bundle", () => {
    it("builds a valid Apps Script bundle", () => {
        const { bundle, sourceHash, sectionCount } = buildCacheFinanceBundle();

        expect(sectionCount).toBe(6);
        expect(sourceHash).toMatch(/^[a-f0-9]{12}$/);
        expect(bundle).toContain(`Source hash: ${sourceHash}`);
        expect(bundle).toContain("function CACHEFINANCE");
        expect(bundle).toContain("class CacheFinanceUtils");
        expect(() => validateBundleContent(bundle)).not.toThrow();
    });

    it("writes dist/CacheFinance.js only when content changes", () => {
        const first = writeCacheFinanceBundle();
        const second = writeCacheFinanceBundle();

        expect(fs.existsSync(paths.dist)).toBe(true);
        expect(first.bundle).toBe(second.bundle);
        expect(isCacheFinanceBundleCurrent()).toBe(true);
    });
});
