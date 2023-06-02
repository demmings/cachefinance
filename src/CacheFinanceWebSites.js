/*  *** DEBUG START ***
//  Remove comments for testing in NODE

export { CacheFinanceWebSites };
export { StockAttributes };
export { FinanceWebSite };
export { TdMarketResearch, GlobeAndMail, YahooFinance };

class Logger {
    static log(msg) {
        console.log(msg);
    }
}
//  *** DEBUG END ***/

/**
 * @classdesc Finance Website Info.
 */
class CacheFinanceWebSites {
    /**
     * All finance website lookup objects should be defined here.
     * All finance Website Objects must implement the method "getInfo(symbol)"
     * The getInfo() method must return an instance of "StockAttributes"
     */
    constructor() {
        this.siteList = [
            new FinanceWebSite("TDEtf", TdMarketsEtf),
            new FinanceWebSite("TDStock", TdMarketsStock),
            new FinanceWebSite("Yahoo", YahooFinance),
            new FinanceWebSite("Globe", GlobeAndMail)
        ];

        /** @property {Map<String, FinanceWebSite>} */
        this.siteMap = new Map();
        for (const site of this.siteList) {
            this.siteMap.set(site.siteName, site.siteObject);
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
        const siteInfo = this.siteMap.get(name);

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
            default:
                countryCode = "ca";     //  We the north!
                break;
        }

        return countryCode;
    }
}

/**
 * @classdesc - defines an instance of an object to perform getInfo(symbol)
 */
class FinanceWebSite {
    /**
     * 
     * @param {String} siteName 
     * @param {object} siteObject - Points to class for getting finance data.
     */
    constructor(siteName, siteObject) {
        this.siteName = siteName;
        this.siteObject = siteObject;
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
 * @classdesc TD Markets lookup by ETF symbol
 */
class TdMarketsEtf {
    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        return TdMarketResearch.getInfo(symbol, "ETF");
    }
}

/**
 * @classdesc TD Markets lookup by STOCK symbol
 */
class TdMarketsStock {
    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        return TdMarketResearch.getInfo(symbol, "STOCK");
    }
}

/**
 * @classdesc Base class to Lookup for TD website.
 */
class TdMarketResearch {
    /**
     * 
     * @param {String} symbol 
     * @param {String} type 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, type = "ETF") {
        const data = new StockAttributes();

        let URL = null;
        if (type === "ETF")
            URL = `https://marketsandresearch.td.com/tdwca/Public/ETFsProfile/Summary/${CacheFinanceWebSites.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;
        else
            URL = `https://marketsandresearch.td.com/tdwca/Public/Stocks/Overview/${CacheFinanceWebSites.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return data;
        }
        Logger.log(`getStockDividendYield:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

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
        const data = new StockAttributes();

        const URL = `https://finance.yahoo.com/quote/${YahooFinance.getTicker(symbol)}`;

        const html = UrlFetchApp.fetch(URL).getContentText();
        Logger.log(`getStockDividendYield:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        let dividendPercent = html.match(/"DIVIDEND_AND_YIELD-value">\d*\.\d*\s\((\d*\.\d*)%\)/);
        if (dividendPercent === null) {
            dividendPercent = html.match(/TD_YIELD-value">(\d*\.\d*)%/);
        }

        if (dividendPercent !== null && dividendPercent.length === 2) {
            const tempPct = dividendPercent[1];
            Logger.log(`PERCENT=${tempPct}`);

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
            Logger.log(`PRICE=${tempPrice}`);

            data.stockPrice = parseFloat(tempPrice);

            if (isNaN(data.stockPrice)) {
                data.stockPrice = null;
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

            modifiedSymbol = symbolParts[1];
            if (symbolParts[0] === "TSE")
                modifiedSymbol = `${symbolParts[1]}.TO`;

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
        const data = new StockAttributes();
        const URL = `https://www.theglobeandmail.com/investing/markets/stocks/${GlobeAndMail.getTicker(symbol)}`;

        Logger.log(`getStockDividendYield:  ${symbol}`);
        Logger.log(`URL = ${URL}`);

        let html = null;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return data;
        }

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
 * Get/Set data about stocks/ETFs.
 */
class StockAttributes {
    constructor() {
        this.yieldPct = null;
        this.stockName = null;
        this.stockPrice = null;
    }

    /**
     * 
     * @param {String} attribute 
     * @returns {any}
     */
    getValue(attribute) {
        switch (attribute) {
            case "PRICE":
                return (this.stockPrice === null) ? 0 : this.stockPrice;

            case "YIELDPCT":
                return (this.yieldPct === null) ? 0 : this.yieldPct;

            case "NAME":
                return (this.stockName === null) ? "" : this.stockName;

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
                return this.stockPrice !== null;

            case "YIELDPCT":
                return this.yieldPct !== null;

            case "NAME":
                return this.stockName !== null;

            default:
                return false;
        }
    }
}
