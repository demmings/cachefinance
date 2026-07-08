#!/usr/bin/env node
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
    isCacheFinanceBundleCurrent,
    paths,
    SOURCE_FILES,
    writeCacheFinanceBundle
} from "./gas-source.mjs";

const args = new Set(process.argv.slice(2));
const checkOnly = args.has("--check");

function formatBytes(bytes) {
    return `${bytes.toLocaleString()} bytes`;
}

function main() {
    if (checkOnly) {
        if (isCacheFinanceBundleCurrent()) {
            console.log(`dist/CacheFinance.js is up to date (${SOURCE_FILES.length} source files)`);
            return;
        }

        console.error("dist/CacheFinance.js is out of date. Run: npm run build");
        process.exit(1);
    }

    const result = writeCacheFinanceBundle();

    if (result.written) {
        console.log(
            `Built ${paths.dist} (${formatBytes(result.bytes)}, ` +
            `${result.sectionCount} files, hash ${result.sourceHash})`
        );
        return;
    }

    console.log(
        `dist/CacheFinance.js is up to date (${formatBytes(result.bytes)}, hash ${result.sourceHash})`
    );
}

const scriptPath = path.resolve(fileURLToPath(import.meta.url));

if (process.argv[1] && path.resolve(process.argv[1]) === scriptPath) {
    main();
}
