/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { SiteThrottle, ThresholdPeriod } from "./CacheFinanceUtils.js";
export { FinanceWebSites };
export { StockAttributes };
export { FinanceWebSite };
export { TdMarketResearch, GlobeAndMail, YahooFinance, YahooApi, FinnHub, AlphaVantage, GoogleWebSiteFinance, TwelveData };

class Logger {
    static log(msg) {
        console.log(msg);
    }
}
//  *** DEBUG END ***/

/**
 * @classdesc Concrete implementations for each finance website access.
 */
class FinanceWebSites {
    /**
     * All finance website lookup objects should be defined here.
     * All finance Website Objects must implement the method "getInfo(symbol)" 
     * The getInfo() method must return an instance of "StockAttributes"
     */
    constructor() {
        this.siteList = [
            new FinanceWebSite("YahooApi", YahooApi),
            new FinanceWebSite("GoogleWebSiteFinance", GoogleWebSiteFinance),
            new FinanceWebSite("FinnHub", FinnHub),
            new FinanceWebSite("Globe", GlobeAndMail),
            new FinanceWebSite("Yahoo", YahooFinance),
            new FinanceWebSite("TwelveData", TwelveData),
            new FinanceWebSite("AlphaVantage", AlphaVantage)
        ];

        /** @property {Map<String, FinanceWebSite>} */
        this.siteMap = new Map();
        for (const site of this.siteList) {
            this.siteMap.set(site.siteName, site);
        }
    }

    /**
     * 
     * @returns {FinanceWebSite[]}
     */
    get() {
        return this.siteList;
    }

    /**
     * Get info for running function to get info about a stock using a specfic service name
     * @param {String} name - Name of search to find object containing class info to make function call.
     * @returns {FinanceWebSite}
     */
    getByName(name) {
        const siteInfo = this.siteMap.get(name.toUpperCase());

        return (typeof siteInfo === 'undefined') ? null : siteInfo;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static getTickerCountryCode(symbol) {
        let exchange = "";
        let countryCode = "";

        if (symbol.indexOf(":") > 0) {
            const parts = symbol.split(":");
            exchange = parts[0].toUpperCase();
        }

        switch (exchange) {
            case "NASDAQ":
            case "NYSEARCA":
            case "NYSE":
            case "NYSEAMERICAN":
            case "OPRA":
            case "OTCMKTS":
            case "BATS":
                countryCode = "us";
                break;
            case "CVE":
            case "TSE":
            case "TSX":
            case "TSXV":
            case "NEO":
                countryCode = "ca";
                break;
            case "SGX":
                countryCode = "sg";
                break;
            case "CURRENCY":
                countryCode = "fx";
                break;
            case "CF":      //  a Mutual fund
                countryCode = "mut";
                break;
            default:
                countryCode = "ca";     //  We the north!
                break;
        }

        return countryCode;
    }

    /**
    * 
    * @param {String} symbol 
    * @returns {String}
    */
    static getBaseTicker(symbol) {
        let modifiedSymbol = symbol;
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const symbolParts = symbol.split(":");

            modifiedSymbol = symbolParts[1];
        }
        return modifiedSymbol;
    }

    /**
     * Get script property using key.  If not found, returns null.
     * @param {String} propertyKey 
     * @returns {any}
     */
    static getApiKey(propertyKey) {
        const scriptProperties = PropertiesService.getScriptProperties();

        const myData = scriptProperties.getProperty(propertyKey);

        return myData;
    }
}

/**
 * @classdesc - defines an instance of an object to perform getInfo(symbol)
 */
class FinanceWebSite {
    /**
     * 
     * @param {String} siteName 
     * @param {object} webSiteJsClass - Points to class for getting finance data.
     */
    constructor(siteName, webSiteJsClass) {
        this.siteName = siteName.toUpperCase();
        this._siteObject = webSiteJsClass;
    }

    set siteName(siteName) {
        this._siteName = siteName;
    }
    get siteName() {
        return this._siteName;
    }

    set siteObject(siteObject) {
        this._siteObject = siteObject;
    }
    get siteObject() {
        return this._siteObject;
    }
}

/**
 * Get/Set data about stocks/ETFs.
 */
class StockAttributes {
    constructor() {
        this._yieldPct = null;
        this._stockName = null;
        this._stockPrice = null;
    }

    get yieldPct() {
        return this._yieldPct;
    }
    set yieldPct(value) {
        if (value !== null) {
            this._yieldPct = Math.round(value * 10000) / 10000;
        }
    }

    get stockPrice() {
        return this._stockPrice;
    }
    set stockPrice(value) {
        if (value !== null) {
            this._stockPrice = Math.round(value * 100) / 100;
        }
    }

    set exchangeRate(value) {
        //  We need 4 decimal places for currency, so setting stockPrice was 
        //  rounding to 2 places - so we have a 'special' price setter for exchange rates.
        if (value !== null) {
            this._stockPrice = Math.round(value * 10000) / 10000;
        }
    }
    get exchangeRate() {
        return this._stockPrice;
    }

    get stockName() {
        return this._stockName;
    }
    set stockName(value) {
        this._stockName = value;
    }

    /**
     * 
     * @param {String} attribute 
     * @returns {any}
     */
    getValue(attribute) {
        switch (attribute) {
            case "PRICE":
                return (this.stockPrice === null) ? '' : this.stockPrice;

            case "YIELDPCT":
                return (this.yieldPct === null) ? '' : this.yieldPct;

            case "NAME":
                return (this.stockName === null) ? '' : this.stockName;

            default:
                return '#N/A';
        }
    }

    /**
     * 
     * @param {String} attribute 
     * @returns {Boolean}
     */
    isAttributeSet(attribute) {
        switch (attribute) {
            case "PRICE":
                {
                    const retVal = this.stockPrice !== null && !isNaN(this.stockPrice) && this.stockPrice !== 0;
                    Logger.log(`price=${this.stockPrice}. Is Valid=${retVal}`);
                    return retVal;
                }

            case "YIELDPCT":
                return this.yieldPct !== null;

            case "NAME":
                return this.stockName !== null;

            default:
                return false;
        }
    }
}


/**
 * @classdesc Lookup for Yahoo site.
 */
class YahooFinance {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        const URL = YahooFinance.getURL(symbol, attribute);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            Logger.log(`FAILED -> getInfo:  ${symbol}.  URL = ${URL}. Err=${ex.toString()}`);
            return new StockAttributes();
        }

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        return YahooFinance.parseResponse(html, symbol);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
        if (FinanceWebSites.getTickerCountryCode(symbol) === "fx") {
            return "";
        }

        return `https://finance.yahoo.com/quote/${YahooFinance.getTicker(symbol)}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return "";
    }

    /**
     * 
     * @param {String} html 
     * @param {String} symbol
     * @returns {StockAttributes}
     */
    static parseResponse(html, symbol) {
        const data = new StockAttributes();

        if (symbol === '') {
            return data;
        }

        //  skipcq:  JS-0097
        let dividendPercent = html.match('title="Yield">Yield</span> <[^>]*>(\\d{0,5}\.?\\d{0,4})');
        if (dividendPercent !== null && dividendPercent.length === 2) {
            const tempPct = dividendPercent[1];
            Logger.log(`Yahoo. Stock=${symbol}. PERCENT=${tempPct}`);

            data.yieldPct = parseFloat(tempPct) / 100;

            if (isNaN(data.yieldPct)) {
                data.yieldPct = null;
            }
        }

        //  skipcq:  JS-0097
        const priceMatch = html.match('qsp-price">(\\d{0,5}\.?\\d{0,4})');
        if (priceMatch !== null && priceMatch.length === 2) {
            const tempPrice = priceMatch[1];
            Logger.log(`Yahoo. Stock=${symbol}.PRICE=${tempPrice}`);

            data.stockPrice = parseFloat(tempPrice);

            if (isNaN(data.stockPrice)) {
                data.stockPrice = null;
            }
        }

        const baseSymbol = YahooFinance.getTicker(symbol);
        // skipcq: JS-0097
        const nameRegex = new RegExp(`<title>(.+?)\(${baseSymbol}\)`);
        const nameMatch = html.match(nameRegex);
        if (nameMatch !== null && nameMatch.length > 1) {
            data.stockName = nameMatch[1].endsWith("(") ? nameMatch[1].slice(0, -1) : nameMatch[1];
            Logger.log(`Yahoo. Stock=${symbol}.NAME=${data.stockName}`);
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static getTicker(symbol) {
        let modifiedSymbol = symbol;
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const symbolParts = symbol.split(":");

            modifiedSymbol = symbolParts[1];
            if (symbolParts[0] === "TSE")
                modifiedSymbol = `${symbolParts[1]}.TO`;
            if (symbolParts[0] === "SGX")
                modifiedSymbol = `${symbolParts[1]}.SI`;
            if (symbolParts[0] === "NEO")
                modifiedSymbol = `${symbolParts[1]}.NE`;

        }
        return modifiedSymbol;
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        return null;
    }
}

class YahooApi {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        const URL = YahooApi.getURL(symbol, attribute);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        return YahooApi.parseResponse(html, symbol);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {String}
     */
    static getURL(symbol, attribute) {
        if (FinanceWebSites.getTickerCountryCode(symbol) === "fx") {
            return "";
        }

        if (attribute !== "PRICE" && attribute !== "NAME") {
            return "";
        }

        return `https://query1.finance.yahoo.com/v8/finance/chart/${YahooApi.getTicker(symbol)}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return "";
    }

    /**
     * 
     * @param {String} html 
     * @param {String} symbol
     * @returns {StockAttributes}
     */
    static parseResponse(html, symbol) {
        const stockData = new StockAttributes();

        if (symbol === '') {
            return stockData;
        }

        try {
            const data = JSON.parse(html);
            if (data?.chart?.result.length > 0) {
                const regularMarketPrice = data.chart.result[0].meta.regularMarketPrice;

                stockData.stockPrice = parseFloat(regularMarketPrice);

                if (isNaN(stockData.stockPrice)) {
                    stockData.stockPrice = null;
                }

                stockData.stockName = data.chart.result[0].meta.longName;
            }
        }
        catch (ex) {
            Logger.log(`Failed to parse JSON: ${symbol}`);
        }

        return stockData;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static getTicker(symbol) {
        return YahooFinance.getTicker(symbol);
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        return null;
    }
}

/**
 * @classdesc Lookup for Globe and Mail website.
 */
class GlobeAndMail {
    /**
     * Only gets dividend yield.
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        const URL = GlobeAndMail.getURL(symbol, attribute);

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }

        return GlobeAndMail.parseResponse(html, symbol);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
        if (FinanceWebSites.getTickerCountryCode(symbol) === "fx") {
            return "";
        }

        //  Mutual fund
        if (FinanceWebSites.getTickerCountryCode(symbol) === "mut") {
            return `https://www.theglobeandmail.com/investing/markets/funds/${GlobeAndMail.getTicker(symbol)}`;
        }

        return `https://www.theglobeandmail.com/investing/markets/stocks/${GlobeAndMail.getTicker(symbol)}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return "";
    }

    /**
     * 
     * @param {String} html 
     * @param {String} symbol
     * @returns {StockAttributes}
     */
    static parseResponse(html, symbol) {
        const data = new StockAttributes();

        //  Get the dividend yield.
        let parts = html.match(/.name="dividendYieldTrailing".*?value="(\d{0,4}\.?\d{0,4})%/);

        if (parts === null) {
            parts = html.match(/.name=\\"dividendYieldTrailing\\".*?value=\\"(\d{0,4}\.?\d{0,4})%/);
        }

        if (parts === null) {
            parts = html.match(/"distributionYield" type="percent" value="(\d{0,4}\.?\d{0,4})%"/);
        }

        if (parts !== null && parts.length === 2) {
            const tempPct = parts[1];

            const parsedValue = parseFloat(tempPct) / 100;

            if (!isNaN(parsedValue)) {
                data.yieldPct = parsedValue;
            }
        }

        //  Get the name.
        parts = html.match(/."symbolName":"(.*?)"/);
        if (parts !== null && parts.length === 2) {
            data.stockName = parts[1];
        }


        //  Get the price.
        parts = html.match(/"lastPrice" value="(\d{0,4}\.?\d{0,4})"/);

        if (parts !== null && parts.length === 2) {

            const parsedValue = parseFloat(parts[1]);
            if (!isNaN(parsedValue)) {
                data.stockPrice = parsedValue;
            }
        }

        return data;
    }

    /**
     * Clean up ticker symbol for use in Globe and Mail lookups.
     * @param {String} symbol 
     * @returns {String}
     */
    static getTicker(symbol) {
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const parts = symbol.split(":");

            switch (parts[0].toUpperCase()) {
                case "TSE":
                    symbol = parts[1];
                    if (parts[1].indexOf(".") !== -1) {
                        symbol = parts[1].replace(".", "-");
                    }
                    else if (parts[1].indexOf("-") !== -1) {
                        const prefShare = parts[1].split("-");
                        symbol = `${prefShare[0]}-PR-${prefShare[1]}`;
                    }
                    symbol = `${symbol}-T`;
                    break;
                case "NEO":
                    symbol = `${parts[1]}-NE`;
                    break;
                case "NYSEARCA":
                    symbol = `${parts[1]}-A`;
                    break;
                case "NASDAQ":
                    symbol = `${parts[1]}-Q`;
                    break;
                case "CF":
                    symbol = `${parts[1]}.CF`;
                    break;
                default:
                    symbol = '#N/A';
            }
        }

        return symbol;
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        return null;
    }
}



/**
 * @classdesc Uses FINNHUB Rest API.  Requires a script setting for the API key.
 * Set key name as FINNHUB_API_KEY
 */
class FinnHub {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns 
     */
    static getInfo(symbol, attribute = "PRICE") {
        let data = new StockAttributes();

        if (attribute !== "PRICE") {
            Logger.log(`Finnhub.  Only PRICE is supported: ${symbol}, ${attribute}`);
            return data;
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (countryCode !== "us") {
            Logger.log(`FinnHub --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const URL = FinnHub.getURL(symbol, attribute, FinnHub.getApiKey());
        Logger.log(`getInfo:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();
            data = FinnHub.parseResponse(jsonStr, symbol, attribute);
        }
        catch (ex) {
            return data;
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @param {String} API_KEY
     * @returns {String}
     */
    static getURL(symbol, attribute, API_KEY = null) {
        if (attribute !== "PRICE") {
            return "";
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (countryCode !== "us") {
            return "";
        }

        if (API_KEY === null) {
            return "";
        }

        return `https://finnhub.io/api/v1/quote?symbol=${FinanceWebSites.getBaseTicker(symbol)}&token=${API_KEY}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return FinanceWebSites.getApiKey("FINNHUB_API_KEY");
    }

    /**
     * 
     * @param {String} jsonStr 
     * @param {String} _symbol
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static parseResponse(jsonStr, _symbol, attribute) {
        const data = new StockAttributes();

        const hubData = JSON.parse(jsonStr);
        if (attribute === "PRICE")
            data.stockPrice = hubData.c;

        return data;
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        //  Basic throttle check
        const limits = [
            new ThresholdPeriod("SECOND", 30),
            new ThresholdPeriod("MINUTE", 60)
        ];

        return new SiteThrottle("FINNHUB", limits);
    }
}

class AlphaVantage {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute = "PRICE") {
        let data = new StockAttributes();

        if (attribute !== "PRICE") {
            Logger.log(`AlphaVantage.  Only PRICE is supported: ${symbol}, ${attribute}`);
            return data;
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (!(countryCode === "us" || countryCode === "fx")) {
            Logger.log(`AlphaVantage --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const apiKey = AlphaVantage.getApiKey();
        const URL = AlphaVantage.getURL(symbol, attribute, apiKey);
        Logger.log(`getInfo: AlphaVantage  ${symbol}.  URL = ${URL}.  Key = ${apiKey}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();
            data = AlphaVantage.parseResponse(jsonStr, symbol, attribute);
        }
        catch (ex) {
            return data;
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @param {String} API_KEY
     * @returns {String}
     */
    static getURL(symbol, attribute, API_KEY = null) {
        if (API_KEY === null) {
            return "";
        }

        if (attribute !== "PRICE") {
            return "";
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (!(countryCode === "us" || countryCode === "fx")) {
            return "";
        }
        const symbolParts = symbol.split(":");

        if (countryCode === "fx") {
            const fromCurrency = symbolParts[1].substring(0, 3);
            const toCurrency = symbolParts[1].substring(3, 6);

            return `https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=${fromCurrency}&to_currency=${toCurrency}&apikey=${API_KEY}`;
        }
        else {
            return `https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=${FinanceWebSites.getBaseTicker(symbol)}&apikey=${API_KEY}`;
        }
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return FinanceWebSites.getApiKey("ALPHA_VANTAGE_API_KEY");
    }

    /**
     * 
     * @param {String} jsonStr 
     * @param {String} symbol
     * @param {String} _attribute
     * @returns {StockAttributes}
     */
    static parseResponse(jsonStr, symbol, _attribute) {
        Logger.log(`content=${jsonStr}`);

        const data = new StockAttributes();
        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        const alphaVantageData = JSON.parse(jsonStr);

        try {
            if (countryCode === "fx") {
                data.exchangeRate = alphaVantageData["Realtime Currency Exchange Rate"]["5. Exchange Rate"];
            }
            else {
                data.stockPrice = alphaVantageData["Global Quote"]["05. price"];
            }

            Logger.log(`Price=${data.stockPrice}`);
        }
        catch (ex) {
            Logger.log(`AlphaVantage JSON Parse Error (looking for ${countryCode}. err=${ex}).`);
        }

        return data;
    }

    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        //  Basic throttle check
        const limits = [
            new ThresholdPeriod("DAY", 25)
        ];

        return new SiteThrottle("ALPHAVANTAGE", limits);
    }
}

/**
 * @classdesc Lookup for GOOGLE Finance site.
 * weirdly, GOOGLEFINANCE() will fail for some stocks, but work on the website.
 */
class GoogleWebSiteFinance {
    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        const URL = GoogleWebSiteFinance.getURL(symbol, attribute);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        return GoogleWebSiteFinance.parseResponse(html, symbol);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {String}
     */
    static getURL(symbol, attribute) {
        if (attribute === "YIELDPCT" && FinanceWebSites.getTickerCountryCode(symbol) === "us") {
            //  This site is very bad at yields for u.s.
            return "";
        }

        return `https://www.google.com/finance/quote/${GoogleWebSiteFinance.getTicker(symbol)}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return "";
    }

    /**
     * 
     * @param {String} html 
     * @param {String} symbol
     * @returns {StockAttributes}
     */
    static parseResponse(html, symbol) {
        const data = new StockAttributes();

        if (symbol === '') {
            return data;
        }

        if (html.indexOf("We couldn't find any match for your search.") !== -1) {
            Logger.log(`www.google.com/finance:  We couldn't find any match for your search. symbol=${symbol}`);
            return data;
        }

        data.yieldPct = GoogleWebSiteFinance.extractYieldPct(html, symbol);
        data.stockName = GoogleWebSiteFinance.extractStockName(html, symbol);
        if (FinanceWebSites.getTickerCountryCode(symbol) === "fx") {
            data.exchangeRate = GoogleWebSiteFinance.extractStockPrice(html, symbol);
        }
        else {
            data.stockPrice = GoogleWebSiteFinance.extractStockPrice(html, symbol);
        }

        Logger.log(`Google. Stock=${symbol}. PERCENT=${data.yieldPct}. NAME=${data.stockName}. PRICE=${data.stockPrice}`);

        return data;
    }

    /**
     * 
     * @param {String} html 
     * @param {String} _symbol 
     * @returns {Number}
     */
    static extractYieldPct(html, _symbol) {
        let data = null;
        //  skipcq: JS-0097
        const dividendPercent = html.match(/Dividend yield.+?(\d{0,4}\.?\d{0,4})%<\/div>/);

        if (dividendPercent !== null && dividendPercent.length > 1) {
            const tempPct = dividendPercent[1];

            data = parseFloat(tempPct) / 100;

            if (isNaN(data)) {
                data = null;
            }
        }
        return data;
    }

    /**
     * 
     * @param {String} html 
     * @param {String} _symbol 
     * @returns {Number}
     */
    static extractStockPrice(html, _symbol) {
        let data = null;
        //  skipcq: JS-0097
        const priceMatch = html.match(/data-last-price="(\d{0,7}\.*\d{0,20})"/);
        if (priceMatch !== null && priceMatch.length > 1) {
            const tempPrice = priceMatch[1];

            data = parseFloat(tempPrice);

            if (isNaN(data)) {
                data = null;
            }
        }

        return data;
    }

    /**
     * 
     * @param {String} html 
     * @param {String} symbol 
     * @returns {String}
     */
    static extractStockName(html, symbol) {
        let data = null;
        const baseSymbol = GoogleWebSiteFinance.getTicker(symbol);
        const stockNameParts = baseSymbol.split(":");

        if (stockNameParts.length > 1) {
            const stock = stockNameParts[0];

            //  skipcq: JS-0097
            const nameRegex = new RegExp(`<title>(.+?)\(${stock}\)`);
            const nameMatch = html.match(nameRegex);
            if (nameMatch !== null && nameMatch.length > 1) {
                data = nameMatch[1].endsWith("(") ? nameMatch[1].slice(0, -1) : nameMatch[1];
            }
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static getTicker(symbol) {
        let modifiedSymbol = symbol;
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const symbolParts = symbol.split(":");

            if (FinanceWebSites.getTickerCountryCode(symbol) !== "fx") {
                modifiedSymbol = `${symbolParts[1]}:${symbolParts[0]}`;
            }
            else {
                const fromCurrency = symbolParts[1].substring(0, 3);
                const toCurrency = symbolParts[1].substring(3, 6);
                modifiedSymbol = `${fromCurrency}-${toCurrency}?hl=en`;
            }
        }
        return modifiedSymbol;
    }


    /**
     * getURL() will receive an instance of the throttling object to query if the limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        return null;
    }
}

class TwelveData {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute = "PRICE") {
        let data = new StockAttributes();

        if (attribute !== "PRICE" && attribute !== "NAME") {
            Logger.log(`TwelveData.  Only PRICE/NAME is supported: ${symbol}, ${attribute}`);
            return data;
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (!(countryCode === "us" || countryCode === "fx")) {
            Logger.log(`TwelveData --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const apiKey = TwelveData.getApiKey();
        const URL = TwelveData.getURL(symbol, attribute, apiKey);
        Logger.log(`getInfo: TwelveData  ${symbol}.  URL = ${URL}.  Key = ${apiKey}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();
            data = TwelveData.parseResponse(jsonStr, symbol, attribute);
        }
        catch (ex) {
            return data;
        }

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @param {String} API_KEY
     * @returns {String}
     */
    static getURL(symbol, attribute, API_KEY = null) {
        if (API_KEY === null) {
            return "";
        }

        if (attribute !== "PRICE" && attribute !== "NAME") {
            return "";
        }

        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        if (!(countryCode === "us" || countryCode === "fx")) {
            return "";
        }

        let twelveDataSymbol = "";
        if (countryCode === "fx") {
            const symbolParts = symbol.split(":");
            const fromCurrency = symbolParts[1].substring(0, 3);
            const toCurrency = symbolParts[1].substring(3, 6);

            twelveDataSymbol = `${fromCurrency}/${toCurrency}`;
        }
        else {
            twelveDataSymbol = FinanceWebSites.getBaseTicker(symbol);
        }

        return `https://api.twelvedata.com/quote?symbol=${twelveDataSymbol}&apikey=${API_KEY}`;
    }

    /**
     * 
     * @returns {String}
     */
    static getApiKey() {
        return FinanceWebSites.getApiKey("TWELVE_DATA_API_KEY");
    }

    /**
     * 
     * @param {String} jsonStr 
     * @param {String} symbol
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static parseResponse(jsonStr, symbol, attribute) {
        Logger.log(`content=${jsonStr}`);

        const data = new StockAttributes();
        const countryCode = FinanceWebSites.getTickerCountryCode(symbol);
        const twelveData = JSON.parse(jsonStr);

        try {
            if (attribute === "NAME") {
                data.stockName = twelveData.name;
                Logger.log(`TwelveData. Name=${data.stockName}`);
            }
            else if (attribute === "PRICE") {
                if (countryCode === "fx") {
                    data.exchangeRate = twelveData.close;
                }
                else {
                    data.stockPrice = twelveData.close;
                }
                Logger.log(`TwelveData. Price=${data.stockPrice}`);
            }
        }
        catch (ex) {
            Logger.log(`TwelveData JSON Parse Error (looking for ${countryCode}. err=${ex}).`);
        }

        return data;
    }

    /**
     * Get an instance of the throttling object to query if the web limit would be exceeded.
     * @returns {SiteThrottle}
     */
    static getThrottleObject() {
        //  Basic throttle check
        const limits = [
            new ThresholdPeriod("MINUTE", 8),
            new ThresholdPeriod("DAY", 800)
        ];

        return new SiteThrottle("TWELVEDATA", limits);
    }
}
