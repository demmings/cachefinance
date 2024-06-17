/*  *** DEBUG START ***
//  Remove comments for testing in NODE

export { FinanceWebSites };
export { StockAttributes };
export { FinanceWebSite };
export { TdMarketResearch, GlobeAndMail, YahooFinance, FinnHub };

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
     * All finance Website Objects must implement the method "getInfo(symbol)" and getPropertyValue(key)
     * The getInfo() method must return an instance of "StockAttributes"
     * The getPropertyValue() is used to query any possible properties known (and unknown future) about the website.
     */
    constructor() {
        this.siteList = [
            new FinanceWebSite("FinnHub", FinnHub),
            new FinanceWebSite("TDEtf", TdMarketsEtf),
            new FinanceWebSite("TDStock", TdMarketsStock),
            new FinanceWebSite("Globe", GlobeAndMail),
            new FinanceWebSite("Yahoo", YahooFinance),
            new FinanceWebSite("AlphaVantage", AlphaVantage),
            new FinanceWebSite("GoogleWebSiteFinance", GoogleWebSiteFinance)
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
                countryCode = "us";
                break;
            case "CVE":
            case "TSE":
            case "TSX":
            case "TSXV":
                countryCode = "ca";
                break;
            case "SGX":
                countryCode = "sg";
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
                return this.stockPrice !== null && this.stockPrice !== 0;

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
 * @classdesc TD Markets lookup by ETF symbol
 */
class TdMarketsEtf {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        return TdMarketResearch.getInfo(symbol, attribute, "ETF");
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
        return TdMarketResearch.getURL(symbol, "ETF");
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
      * @param {String} _symbol
      * @returns {StockAttributes}
      */
    static parseResponse(html, _symbol) {
        return TdMarketResearch.parseResponse(html);
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
    }
}

/**
 * @classdesc TD Markets lookup by STOCK symbol
 */
class TdMarketsStock {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute
     * @returns {StockAttributes}
     */
    static getInfo(symbol, attribute) {
        return TdMarketResearch.getInfo(symbol, attribute, "STOCK");
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
        return TdMarketResearch.getURL(symbol, "STOCK");
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
      * @param {String} _symbol
      * @returns {StockAttributes}
      */
    static parseResponse(html, _symbol) {
        return TdMarketResearch.parseResponse(html);
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
    }
}

/**
 * @classdesc Base class to Lookup for TD website.
 */
class TdMarketResearch {
    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @param {String} type 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, _attribute, type = "ETF") {
        const URL = TdMarketResearch.getURL(symbol, type);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }
        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        return TdMarketResearch.parseResponse(html);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} type 
     * @returns {String}
     */
    static getURL(symbol, type = "ETF") {
        let URL = null;
        if (type === "ETF")
            URL = `https://marketsandresearch.td.com/tdwca/Public/ETFsProfile/Summary/${FinanceWebSites.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;
        else
            URL = `https://marketsandresearch.td.com/tdwca/Public/Stocks/Overview/${FinanceWebSites.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;

        return URL;
    }

    /**
     * 
     * @param {String} html 
     * @returns {StockAttributes}
     */
    static parseResponse(html) {
        const data = new StockAttributes();

        //  Get the dividend yield.

        let parts = html.match(/Dividend Yield<\/th><td class="last">(\d{0,4}\.?\d{0,4})%/);
        if (parts === null) {
            parts = html.match(/Dividend Yield<\/div>.*?cell-container contains">(\d{0,4}\.?\d{0,4})%/);
        }
        if (parts !== null && parts.length === 2) {
            const tempPct = parts[1];

            const parsedValue = parseFloat(tempPct) / 100;

            if (!isNaN(parsedValue)) {
                data.yieldPct = parsedValue;
            }
        }

        //  Get the name.
        parts = html.match(/.<span class="issueName">(.*?)<\//);
        if (parts !== null && parts.length === 2) {
            data.stockName = parts[1];
        }

        //  Get the price.
        parts = html.match(/.LAST PRICE<\/span><div><span>(\d{0,4}\.?\d{0,4})</);
        if (parts !== null && parts.length === 2) {

            const parsedValue = parseFloat(parts[1]);
            if (!isNaN(parsedValue)) {
                data.stockPrice = parsedValue;
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
        const colon = symbol.indexOf(":");

        if (colon >= 0) {
            const parts = symbol.split(":");
            symbol = parts[1];
        }

        const dash = symbol.indexOf("-");
        if (dash >= 0) {
            symbol = symbol.replace("-", ".PR.");
        }

        return symbol;
    }
}

/**
 * @classdesc Lookup for Yahoo site.
 */
class YahooFinance {
    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        const URL = YahooFinance.getURL(symbol);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
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

        let dividendPercent = html.match(/Forward Dividend &amp; Yield.+?\((\d*\.\d*)%\)/);
        if (dividendPercent === null) {
            dividendPercent = html.match(/TD_YIELD-value">(\d*\.\d*)%/);
        }

        if (dividendPercent !== null && dividendPercent.length === 2) {
            const tempPct = dividendPercent[1];
            Logger.log(`Yahoo. Stock=${symbol}. PERCENT=${tempPct}`);

            data.yieldPct = parseFloat(tempPct) / 100;

            if (isNaN(data.yieldPct)) {
                data.yieldPct = null;
            }
        }

        const baseSymbol = YahooFinance.getTicker(symbol);
        const re = new RegExp(`data-symbol="${baseSymbol}".+?"regularMarketPrice".+?value="(\\d*\\.?\\d*)?"`);

        const priceMatch = html.match(re);

        if (priceMatch !== null && priceMatch.length === 2) {
            const tempPrice = priceMatch[1];
            Logger.log(`Yahoo. Stock=${symbol}.PRICE=${tempPrice}`);

            data.stockPrice = parseFloat(tempPrice);

            if (isNaN(data.stockPrice)) {
                data.stockPrice = null;
            }
        }

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
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
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

        }
        return modifiedSymbol;
    }
}

/**
 * @classdesc Lookup for Globe and Mail website.
 */
class GlobeAndMail {
    /**
     * Only gets dividend yield.
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        const URL = GlobeAndMail.getURL(symbol);

        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return new StockAttributes();
        }

        return GlobeAndMail.parseResponse(html);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
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
     * @returns {StockAttributes}
     */
    static parseResponse(html) {
        const data = new StockAttributes();

        //  Get the dividend yield.
        let parts = html.match(/.name="dividendYieldTrailing".*?value="(\d{0,4}\.?\d{0,4})%/);

        if (parts === null)
            parts = html.match(/.name=\\"dividendYieldTrailing\\".*?value=\\"(\d{0,4}\.?\d{0,4})%/);

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
        parts = html.match(/."lastPrice":"(\d{0,4}\.?\d{0,4})"/);
        if (parts !== null && parts.length === 2) {

            const parsedValue = parseFloat(parts[1]);
            if (!isNaN(parsedValue)) {
                data.stockPrice = parsedValue;
            }
        }

        return data;
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
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
                case "NYSEARCA":
                    symbol = `${parts[1]}-A`;
                    break;
                case "NASDAQ":
                    symbol = `${parts[1]}-Q`;
                    break;
                default:
                    symbol = '#N/A';
            }
        }

        return symbol;
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
            data = FinnHub.parseResponse(jsonStr);
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
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
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
        if (countryCode !== "us") {
            Logger.log(`AlphaVantage --> Only U.S. stocks: ${symbol}`);
            return data;
        }

        const URL = AlphaVantage.getURL(symbol, attribute, AlphaVantage.getApiKey());
        Logger.log(`getInfo:  ${symbol}.  URL = ${URL}`);

        let jsonStr = null;
        try {
            jsonStr = UrlFetchApp.fetch(URL).getContentText();
            data = AlphaVantage.parseResponse(jsonStr);
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
        if (countryCode !== "us") {
            return "";
        }

        return `https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=${FinanceWebSites.getBaseTicker(symbol)}&apikey=${API_KEY}`;
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
     * @returns {StockAttributes}
     */
    static parseResponse(jsonStr) {
        const data = new StockAttributes();

        Logger.log(`content=${jsonStr}`);
        try {
            const alphaVantageData = JSON.parse(jsonStr);
            data.stockPrice = alphaVantageData["Global Quote"]["05. price"];
            Logger.log(`Price=${data.stockPrice}`);
        }
        catch (ex) {
            Logger.log("AlphaVantage JSON Parse Error.");
        }

        return data;
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
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
    static getInfo(symbol) {
        const URL = GoogleWebSiteFinance.getURL(symbol);

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
     * @param {String} _attribute
     * @returns {String}
     */
    static getURL(symbol, _attribute) {
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

        data.yieldPct = GoogleWebSiteFinance.extractYieldPct(html, symbol);
        data.stockPrice = GoogleWebSiteFinance.extractStockPrice(html, symbol);
        data.stockName = GoogleWebSiteFinance.extractStockName(html, symbol);

        return data;
    }

    static extractYieldPct(html, symbol) {
        let data = null;
        const divReg = new RegExp("Dividend yield.+?(\d+([.]\d*)?|[.]\d+)%<\/div>");
        const dividendPercent = html.match(divReg);

        if (dividendPercent !== null && dividendPercent.length > 1) {
            const tempPct = dividendPercent[1];
            Logger.log(`Google. Stock=${symbol}. PERCENT=${tempPct}`);

            data = parseFloat(tempPct) / 100;

            if (isNaN(data)) {
                data = null;
            }
        }
        return data;
    }

    static extractStockPrice(html, symbol) {
        let data = null;
        const re = new RegExp('data-last-price="(\\d*\\.?\\d*)?"');

        const priceMatch = html.match(re);

        if (priceMatch !== null && priceMatch.length === 2) {
            const tempPrice = priceMatch[1];
            Logger.log(`Google. Stock=${symbol}.PRICE=${tempPrice}`);

            data = parseFloat(tempPrice);

            if (isNaN(data)) {
                data = null;
            }
        }

        return data;
    }

    static extractStockName(html, symbol) {
        let data = null;
        const baseSymbol = GoogleWebSiteFinance.getTicker(symbol);
        const stockNameParts = baseSymbol.split(":");

        if (stockNameParts.length > 1) {
            const stock = stockNameParts[0];

            const nameRegex = new RegExp(`<title>(.+?)\(${stock}\)`);
            const nameMatch = html.match(nameRegex);
            if (nameMatch !== null && nameMatch.length > 1) {
                data = nameMatch[1].endsWith("(") ? nameMatch[1].slice(0, -1) : nameMatch[1];
                Logger.log(`Google. Stock=${symbol}.NAME=${data.stockName}`);
            }
        }

        return data;
    }

    /**
     * 
     * @param {String} key 
     * @param {any} defaultValue 
     * @returns {any}
     */
    static getPropertyValue(key, defaultValue) {
        return defaultValue;
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

            modifiedSymbol = symbolParts[1] + ":" + symbolParts[0];
        }
        return modifiedSymbol;
    }
}