import { defineConfig } from "vitest/config";
import { enableDebugExports } from "./scripts/gas-source.mjs";

export default defineConfig({
    test: {
        environment: "node",
        setupFiles: ["./test/setup.js"],
        include: ["test/**/*.test.js"],
        coverage: {
            provider: "v8",
            reporter: ["text", "text-summary", "html", "lcov"],
            reportsDirectory: "./coverage",
            include: ["src/**/*.js"],
            exclude: [
                "src/GasMocks.js",
                "src/CacheFinanceTest.js"
            ],
            thresholds: {
                lines: 85,
                functions: 85,
                branches: 80,
                statements: 85
            }
        }
    },
    plugins: [
        {
            name: "cachefinance-gas-debug",
            transform(code, id) {
                if (id.includes("/src/") && id.endsWith(".js")) {
                    return enableDebugExports(code);
                }
            }
        }
    ]
});
