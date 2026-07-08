import { describe, expect, it } from "vitest";
import {
    AlphaVantage,
    CoinMarket,
    FinanceWebSites,
    FinanceWebSite,
    GlobeAndMail,
    GoogleWebSiteFinance,
    TwelveData,
    YahooApi,
    YahooFinance,
    FinnHub,
    StockAttributes
} from "../src/CacheFinanceWebSites.js";
import { PropertiesService } from "../src/GasMocks.js";

describe("FinanceWebSites.getTickerCountryCode", () => {
    it.each([
        ["NASDAQ:AAPL", "us"],
        ["NYSEARCA:VOO", "us"],
        ["TSE:ZTL", "ca"],
        ["NEO:CJP", "ca"],
        ["CURRENCY:USDEUR", "fx"],
        ["CF:DYN2752", "mut"]
    ])("maps %s to %s", (symbol, countryCode) => {
        expect(FinanceWebSites.getTickerCountryCode(symbol)).toBe(countryCode);
    });
});

describe("GlobeAndMail", () => {
    it("formats Canadian tickers for Globe and Mail URLs", () => {
        expect(GlobeAndMail.getTicker("TSE:FTN-A")).toBe("FTN-PR-A-T");
        expect(GlobeAndMail.getTicker("NEO:CJP")).toBe("CJP-NE");
        expect(GlobeAndMail.getTicker("CF:DYN2752")).toBe("DYN2752.CF");
    });

    it("builds stock and mutual fund URLs", () => {
        expect(GlobeAndMail.getURL("TSE:FTN-A", "PRICE"))
            .toBe("https://www.theglobeandmail.com/investing/markets/stocks/FTN-PR-A-T");
        expect(GlobeAndMail.getURL("CF:DYN2752", "PRICE"))
            .toBe("https://www.theglobeandmail.com/investing/markets/funds/DYN2752.CF");
        expect(GlobeAndMail.getURL("CURRENCY:USDEUR", "PRICE")).toBe("");
    });

    it("parses price, yield, and name from HTML", () => {
        const html = `
            "symbolName":"Canadian Premium Fund"
            "lastPrice" value="12.34"
            name="dividendYieldTrailing" type="percent" value="3.25%"
        `;

        const data = GlobeAndMail.parseResponse(html);
        expect(data.stockName).toBe("Canadian Premium Fund");
        expect(data.stockPrice).toBe(12.34);
        expect(data.yieldPct).toBeCloseTo(0.0325);
    });
});

describe("YahooFinance", () => {
    it("translates exchange codes into Yahoo ticker format", () => {
        expect(YahooFinance.getTicker("TSE:ZTL")).toBe("ZTL.TO");
        expect(YahooFinance.getTicker("NEO:CJP")).toBe("CJP.NE");
    });

    it("parses price, yield, and name from HTML", () => {
        const html = `
            <title>Vanguard S&P 500 ETF(VOO)
            <span title="Yield">Yield</span> <span>1.42</span>
            <fin-streamer data-field="regularMarketPrice" data-symbol="VOO" data-value="412.34" class="qsp-price">412.34</fin-streamer>
        `;

        const data = YahooFinance.parseResponse(html, "NYSEARCA:VOO");
        expect(data.stockName).toBe("Vanguard S&P 500 ETF");
        expect(data.yieldPct).toBeCloseTo(0.0142);
        expect(data.stockPrice).toBe(412.34);
    });
});

describe("FinnHub", () => {
    it("returns a quote URL only for supported US price lookups", () => {
        expect(FinnHub.getURL("NYSEARCA:VOO", "PRICE", "test-key"))
            .toBe("https://finnhub.io/api/v1/quote?symbol=VOO&token=test-key");
        expect(FinnHub.getURL("TSE:ZTL", "PRICE", "test-key")).toBe("");
        expect(FinnHub.getURL("NYSEARCA:VOO", "NAME", "test-key")).toBe("");
    });

    it("parses Finnhub quote JSON", () => {
        const data = FinnHub.parseResponse('{"c": 123.456}', "NYSEARCA:VOO", "PRICE");
        expect(data.stockPrice).toBe(123.46);
    });
});

describe("StockAttributes", () => {
    it("rounds stock prices to two decimal places", () => {
        const data = new StockAttributes();
        data.stockPrice = 12.3456;
        expect(data.stockPrice).toBe(12.35);
    });

    it("rounds exchange rates to four decimal places", () => {
        const data = new StockAttributes();
        data.exchangeRate = 1.234567;
        expect(data.exchangeRate).toBe(1.2346);
    });

    it("returns attribute values and detects whether they are set", () => {
        const data = new StockAttributes();
        data.stockPrice = 10;
        data.yieldPct = 0.02;
        data.stockName = "Example";

        expect(data.getValue("PRICE")).toBe(10);
        expect(data.getValue("YIELDPCT")).toBe(0.02);
        expect(data.getValue("NAME")).toBe("Example");
        expect(data.isAttributeSet("PRICE")).toBe(true);
        expect(data.isAttributeSet("YIELDPCT")).toBe(true);
        expect(data.isAttributeSet("LOW52")).toBe(false);
    });
});

describe("FinanceWebSites helpers", () => {
    it("extracts base tickers and currency pairs", () => {
        expect(FinanceWebSites.getBaseTicker("NYSEARCA:VOO")).toBe("VOO");
        expect(FinanceWebSites.getCurrencyTickers("CURRENCY:USDEUR")).toEqual({
            fromCurrency: "USD",
            toCurrency: "EUR"
        });
    });

    it("reads API keys from script properties", () => {
        PropertiesService.getScriptProperties().setProperty("FINNHUB_API_KEY", "secret-key");
        expect(FinanceWebSites.getApiKey("FINNHUB_API_KEY")).toBe("secret-key");
    });
});

describe("AlphaVantage", () => {
    it("builds US and FX quote URLs when an API key is present", () => {
        expect(AlphaVantage.getURL("NYSEARCA:VOO", "PRICE", "abc"))
            .toBe("https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=VOO&apikey=abc");
        expect(AlphaVantage.getURL("CURRENCY:USDEUR", "PRICE", "abc"))
            .toBe("https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=USD&to_currency=EUR&apikey=abc");
    });

    it("parses stock and FX JSON responses", () => {
        const stock = AlphaVantage.parseResponse(
            JSON.stringify({ "Global Quote": { "05. price": "123.45" } }),
            "NYSEARCA:VOO",
            "PRICE"
        );
        const fx = AlphaVantage.parseResponse(
            JSON.stringify({ "Realtime Currency Exchange Rate": { "5. Exchange Rate": "1.2345" } }),
            "CURRENCY:USDEUR",
            "PRICE"
        );

        expect(stock.stockPrice).toBe(123.45);
        expect(fx.exchangeRate).toBe(1.2345);
    });
});

describe("TwelveData", () => {
    it("builds quote URLs for stocks and currencies", () => {
        PropertiesService.getScriptProperties().setProperty("TWELVE_DATA_API_KEY", "td-key");

        expect(TwelveData.getURL("NYSEARCA:VOO", "PRICE", "td-key"))
            .toBe("https://api.twelvedata.com/quote?symbol=VOO&apikey=td-key");
        expect(TwelveData.getURL("CURRENCY:USDEUR", "PRICE", "td-key"))
            .toBe("https://api.twelvedata.com/quote?symbol=USD/EUR&apikey=td-key");
    });

    it("parses name and price fields from JSON", () => {
        const nameData = TwelveData.parseResponse(
            JSON.stringify({ name: "Vanguard ETF" }),
            "NYSEARCA:VOO",
            "NAME"
        );
        const priceData = TwelveData.parseResponse(
            JSON.stringify({ close: "412.3" }),
            "NYSEARCA:VOO",
            "PRICE"
        );

        expect(nameData.stockName).toBe("Vanguard ETF");
        expect(priceData.stockPrice).toBe(412.3);
    });
});

describe("GoogleWebSiteFinance", () => {
    it("skips US yield lookups", () => {
        expect(GoogleWebSiteFinance.getURL("NYSEARCA:VOO", "YIELDPCT")).toBe("");
    });

    it("returns empty data when Google reports no match", () => {
        const data = GoogleWebSiteFinance.parseResponse(
            "We couldn't find any match for your search.",
            "TSE:ZTL"
        );

        expect(data.stockPrice).toBeNull();
        expect(data.stockName).toBeNull();
    });
});

describe("CoinMarket", () => {
    it("builds crypto conversion URLs when an API key is present", () => {
        expect(CoinMarket.getURL("CURRENCY:BTCUSD", "PRICE", "cmc-key"))
            .toBe("https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol=BTC&convert=USD&CMC_PRO_API_KEY=cmc-key");
    });

    it("parses crypto quote JSON", () => {
        const json = JSON.stringify({
            data: {
                BTC: {
                    name: "Bitcoin",
                    quote: {
                        USD: { price: 50000.12 }
                    }
                }
            }
        });

        const data = CoinMarket.parseResponse(json, "CURRENCY:BTCUSD", "PRICE");
        expect(data.stockName).toBe("Bitcoin");
        expect(data.exchangeRate).toBe(50000.12);
    });
});

describe("YahooApi", () => {
    it("builds chart API URLs for supported attributes", () => {
        expect(YahooApi.getURL("NASDAQ:VTC", "PRICE"))
            .toBe("https://query1.finance.yahoo.com/v8/finance/chart/VTC");
        expect(YahooApi.getURL("NASDAQ:VTC", "YIELDPCT")).toBe("");
        expect(YahooApi.getURL("CURRENCY:USDEUR", "PRICE")).toBe("");
    });

    it("parses chart JSON responses", () => {
        const json = JSON.stringify({
            chart: {
                result: [{
                    meta: {
                        regularMarketPrice: 123.456,
                        longName: "Vanguard Tax-Exempt Bond"
                    }
                }]
            }
        });

        const data = YahooApi.parseResponse(json, "NASDAQ:VTC");
        expect(data.stockPrice).toBe(123.46);
        expect(data.stockName).toBe("Vanguard Tax-Exempt Bond");
    });
});

describe("GoogleWebSiteFinance", () => {
    it("formats tickers for stocks and currencies", () => {
        expect(GoogleWebSiteFinance.getTicker("TSE:ZTL")).toBe("ZTL:TSE");
        expect(GoogleWebSiteFinance.getTicker("CURRENCY:USDEUR")).toBe("USD-EUR?hl=en");
    });

    it("extracts yield, price, and name from Google Finance HTML", () => {
        const html = `
            <title>Canadian Premium Fund(ZTL)
            Dividend yield</span>2.15%</div>
            data-last-price="18.76"
        `;

        expect(GoogleWebSiteFinance.extractYieldPct(html, "TSE:ZTL")).toBeCloseTo(0.0215);
        expect(GoogleWebSiteFinance.extractStockPrice(html, "TSE:ZTL")).toBe(18.76);
        expect(GoogleWebSiteFinance.extractStockName(html, "TSE:ZTL")).toBe("Canadian Premium Fund");
    });

    it("parses a complete Google Finance response for stocks", () => {
        const html = `
            <title>Canadian Premium Fund(ZTL)
            Dividend yield</span>2.15%</div>
            data-last-price="18.76"
        `;

        const data = GoogleWebSiteFinance.parseResponse(html, "TSE:ZTL");
        expect(data.stockPrice).toBe(18.76);
        expect(data.yieldPct).toBeCloseTo(0.0215);
        expect(data.stockName).toBe("Canadian Premium Fund");
    });
});

describe("FinanceWebSites registry", () => {
    it("looks up finance site objects by name", () => {
        const sites = new FinanceWebSites();
        const yahoo = sites.getByName("yahoo");

        expect(yahoo).not.toBeNull();
        expect(yahoo.siteName).toBe("YAHOO");
        expect(sites.getByName("missing")).toBeNull();
    });

    it("exposes all configured providers", () => {
        const sites = new FinanceWebSites().get().map((site) => site.siteName);
        expect(sites).toEqual(expect.arrayContaining(["YAHOOAPI", "FINNHUB", "COINMARKET"]));
    });
});

describe("FinanceWebSite", () => {
    it("stores site metadata and parser references", () => {
        const parser = { getInfo: () => new StockAttributes() };
        const site = new FinanceWebSite("Custom", parser);

        expect(site.siteName).toBe("CUSTOM");
        expect(site.siteObject).toBe(parser);

        site.siteName = "Updated";
        site.siteObject = parser;
        expect(site.siteName).toBe("Updated");
    });
});
