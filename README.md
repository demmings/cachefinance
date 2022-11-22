# About

<table>
<tr>
<td>

    * CACHEFINANCE is a custom function and trigger to supplement GOOGLEFINANCE.
    * Valid **STOCK** data is always available even when GOOGLEFINANCE refuses to work.
    * GOOGLEFINANCE does not support all stock symbols.  Unsupported stocks can get price/name/yield data.  
    * As you can guess from the name, data is cached so when '#N/A' appears it does not mess up your asset history logging/graphing.
    
</td>
</tr>
</table>

# Installing

    * Copy files manually.
    * In the ./dist folder there is **ONE** required file:
      * CacheFinance.js  
      * This file is an amalgamation of all files in the **/src** folder
      * None of the files in ./src are required if you use **dist/CacheFinance.js**
    * **OR** in the ./src folder there are **TWO** required files:
      * CacheFinance.js
      * ScriptSettings.js
    * The simple approach is to copy and paste each file.
      * From your sheets Select **Extensions** and then **Apps Script**
      * Ensure that Editor is selected.  It is the **< >**
      * Click the PLUS sign beside **File** and then select **Script**
      * Find each file in turn in the **src** OR **dist** folder in the Github repository.
      * Click on a file, and then click on **Copy Raw Contents** which puts the file into your copy buffer.
      * Back in your Google Project, rename **Untitled** to the file name you just selected in Github.  It is not necessary to enter the .gs extension.
      * Remove the default contents of the file **myFunction()** and paste in the new content you have copied from Github (Ctrl-v).
      * Click the little diskette icon to save.
      * Continue with all five files until done.
      * Change to your spreadsheet screen and try typing in any cell
        * ```=CACHEFINANCE()```.  The new function with online help should be available.


# Using
    * After adding the script, it will require new permissions.
    * You need to open the script inside the Google Script editor, go to the Run menu and choose 'CacheFinanceBoot' from the dropdown. This will prompt you to authorize the script and the triggers will function with the correct permissions.

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
      *  CACHEFINANCE(symbol, attribute, defaultValue)
      *  symbol - stock symbol using regular GOOGLEFINANCE conventions.
      *  attribute - three supporte attributes for now "price", "yieldpct", "name".
      *  defaultValue - Use GOOGLEFINANCE() to supply this value either directly or using a CELL that contains the GOOGLEFINANCE value.
         *  'yieldpct' does not work for STOCKS and ETF's in GOOGLEFINANCE, so don't supply the third parameter when using that attribute.
      *  Example:
         *  ```=CACHEFINANCE("TSE:ZTL", "price", GOOGLEFINANCE("TSE:ZTL", "price"))```

![Trigger Setup](img/CACHEFINANCE_LEGEND.png)
