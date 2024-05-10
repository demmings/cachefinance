![Your Repository’s Stats](https://github-readme-stats.vercel.app/api?username=demmings&show_icons=true)

[![Quality gate](https://sonarcloud.io/api/project_badges/quality_gate?project=demmings_cachefinance)](https://sonarcloud.io/summary/new_code?id=demmings_cachefinance)
[![Code Smells](https://sonarcloud.io/api/project_badges/measure?project=demmings_cachefinance&metric=code_smells)](https://sonarcloud.io/summary/new_code?id=demmings_cachefinance)
[![Maintainability Rating](https://sonarcloud.io/api/project_badges/measure?project=demmings_cachefinance&metric=sqale_rating)](https://sonarcloud.io/summary/new_code?id=demmings_cachefinance)
[![Bugs](https://sonarcloud.io/api/project_badges/measure?project=demmings_cachefinance&metric=bugs)](https://sonarcloud.io/summary/new_code?id=demmings_cachefinance)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=demmings_cachefinance&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=demmings_cachefinance)
[![Security Rating](https://sonarcloud.io/api/project_badges/measure?project=demmings_cachefinance&metric=security_rating)](https://sonarcloud.io/summary/new_code?id=demmings_cachefinance)
[![Vulnerabilities](https://sonarcloud.io/api/project_badges/measure?project=demmings_cachefinance&metric=vulnerabilities)](https://sonarcloud.io/summary/new_code?id=demmings_cachefinance)
[![DeepSource](https://deepsource.io/gh/demmings/cachefinance.svg/?label=active+issues&show_trend=true&token=uIplDc6IW1XQfmDks0l97l4C)](https://deepsource.io/gh/demmings/cachefinance/?ref=repository-badge)


---

# About

<table>
<tr>
<td>

* **CACHEFINANCE** is a custom function to supplement GOOGLEFINANCE.
  * Use this for ONE symbol and ONE attribute lookup.
* * **CACHEFINANCES** is a custom function similar to **CACHEFINANCE** except it is used to process a range of symbols.
* Valid **STOCK** data is always available even when GOOGLEFINANCE refuses to work.
* GOOGLEFINANCE does not support all stock symbols.  Many unsupported google stocks can still get price/name/yield data (using web screen scraping)
* As you can guess from the name, data is cached so when '#N/A' appears, it uses the last known value so that it does not mess up your asset history logging/graphing.
* [All My Google Sheets Work](https://demmings.github.io/index.html)
* [CacheFinance Web site](https://demmings.github.io/notes/cachefinance.html)
    
</td>
</tr>
</table>

# Installing

* Copy files manually.
* In the ./dist folder there are two files.  Only one is required.  Choose only **ONE** of the files based on your needs:
    * **CacheFinance.js**  
      * Caches GOOGLEFINANCE results AND does 3'rd party website lookups when all else fails.
        * This file is an amalgamation of the following files in the **/src** folder
          * CacheFinance.js
          * CacheFinance3rdParty.js
          * CacheFinanceWebSites.js
          * CacheFinanceTest.js
          * ScriptSettings.js
          * CacheFinanceUtils.js
      * None of the files in ./src are required if you use **dist/CacheFinance.js**
    * **CacheFinanceTrigger.js** (deprecated)
      * All of the functionality of **CacheFinance.js**  PLUS  use a trigger to pull your data.
* The simple approach is to copy and paste each file.
    * From your sheets Select **Extensions** and then **Apps Script**
    * Ensure that Editor is selected.  It is the **< >**
    * Click the PLUS sign beside **File** and then select **Script**
    * Find each file in turn in the **src** OR **dist** folder in the Github repository.
    * Click on a file, and then click on **Copy Raw Contents** which puts the file into your copy buffer.
    * Back in your Google Project, rename **Untitled** to the file name you just selected in Github.  It is not necessary to enter the .gs extension.
    * Remove the default contents of the file **myFunction()** and paste in the new content you have copied from Github (Ctrl-v).
    * Click the little diskette icon to save.
    * Continue with all files until done.
    * Change to your spreadsheet screen and try typing in any cell
    * ```=CACHEFINANCE()```.  The new function with online help should be available.


# Using
* After adding the script, it will require new permissions.
* You need to open the script inside the Google Script editor, go to the Run menu and choose 'CacheFinanceBoot' from the dropdown. This will prompt you to authorize the script and the triggers will function with the correct permissions.
* **Finnhub** 
  * For faster U.S. stock price lookups when external finance data is used, add the key to **Apps Script** ==> **Project Settings** ==> **Script Properties**
    * Click on **Edit Script Properties** ==> **Add Script Property**.  
      * Set the property name to:  **FINNHUB_API_KEY**
      * Set the value to:  *'YOUR FINNHUB API KEY'*
        * Get your API key at:  https://finnhub.io/

## Using as a custom function.
* The custom function **CACHEFINANCE** enhances the capabilities of GOOGLEFINANCE.
* When it is working, GOOGLEFINANCE() is much faster to retrieve stock data than calling a URL and scraping the finance data - so it is used as the default source of information.
* When GOOGLEFINANCE() works, the data is cached.
* When GOOGLEFINANCE() fails ('#N/A'), CACHEFINANCE() will search for a cached version of the data.  It is better to return a reasonable value, rather than just fail.  If your asset tracking scripts have just one bad data point, your total values will be invalid.
* If the data cannot be found in cache, the function will attempt to find the data at various financial websites.  This process however can take several seconds just to retrieve one data point.
* If this also fails, PRICE and YIELDPCT return 0, while NAME returns an empty string.
* **CAVEAT EMPTOR**.  Custom functions are also far from perfect.  If Google Sheets decides to throw up the dreaded 'Loading' error, you are almost back to where we started with an unreliable GOOGLEFINANCE() function.
     * However, in my testing it seems to happen more often when you are doing a large number of finance lookups. 
* **SYNTAX**.
    *  ```CACHEFINANCE(symbol, attribute, defaultValue)```
    * **symbol** - stock symbol using regular GOOGLEFINANCE conventions.
    * **attribute** - three supported attributes doing 3'rd party website lookups:  
       * "price" 
       * "yieldpct"
       * "name"
       * "test" -  special case.  Lists in a table results of tests to third party finance sites.
         * ```CACHEFINANCE("", "TEST")```
       * "clearcache" - special case.  Removes **ALL** CACHEFINANCE entries in script settings.  This will force a re-test of all finance websites the next time CACHEFINANCE cannot get valid data from GOOGLEFINANCE.
      * You can specify other attributes that GOOGLEFINANCE uses, but the CacheFinance() function will not look up this data if GOOGLEFINANCE does not provide an initial default value.
      * This ATTRIBUTE name in this case is used to create our CACHE key, so its name is not important - other than when the function does a cache lookup using this key (which is made by **attribute + "|" + symbol**)
      * The following "low52" does not lookup 3'rd party website data, it will just save any value returned by GOOGLEFINANCE to cache, for the case when GOOGLEFINANCE fails to work:
    ```
        =CACHEFINANCE("TSE:ZIC","low52", GOOGLEFINANCE("TSE:ZIC", "low52"))
    ```
    * **defaultValue** - Use GOOGLEFINANCE() to supply this value either directly or using a CELL that contains the GOOGLEFINANCE value.
      * 'yieldpct' does not work for STOCKS and ETF's in GOOGLEFINANCE, so don't supply the third parameter when using that attribute.
    * Example: (symbol that is not recognized by GOOGLEFINANCE)
        *  ```=CACHEFINANCE("TSE:ZTL", "price", GOOGLEFINANCE("TSE:ZTL", "price"))```

## CACHEFINANCES
* **SYNTAX**.
    *  ```CACHEFINANCE(symbolRange, attribute, defaultValueRange, cacheSeconds)```
* **EXAMPLES**
```
=CACHEFINANCES(A30:A164, "price", B30:B164, 1200)
```
  * symbol range:  A30:A164
  * attribute:     "price"
  * default range (provided by GOOGLEFINANCE) : B30:B164
  * cache seconds: 1200  

```
=CACHEFINANCES(A30:A164, "Yieldpct",,21599)
```
  * symbol range:  A30:A164
  * attribute:     "YieldPct"
  * default values: Not used (leave empty)
  * cache seconds:  21599  (for data points that don't change often, make this a higher number - max=21600 seconds)


## Using through a trigger. (deprecated)
* [cache finance trigger](CacheFinanceTrigger.md)
