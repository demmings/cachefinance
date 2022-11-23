![Your Repository’s Stats](https://github-readme-stats.vercel.app/api?username=demmings&show_icons=true)

---

# About

<table>
<tr>
<td>

* **CACHEFINANCE** is a custom function and trigger to supplement GOOGLEFINANCE.
* Valid **STOCK** data is always available even when GOOGLEFINANCE refuses to work.
* GOOGLEFINANCE does not support all stock symbols.  Many unsupported google stocks can still get price/name/yield data (using web screen scraping)
* As you can guess from the name, data is cached so when '#N/A' appears, it uses the last known value so that it does not mess up your asset history logging/graphing.
    
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
    * Continue with all files until done.
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
    *  ```CACHEFINANCE(symbol, attribute, defaultValue)```
    *  symbol - stock symbol using regular GOOGLEFINANCE conventions.
    *  attribute - three supporte attributes for now "price", "yieldpct", "name".
    *  defaultValue - Use GOOGLEFINANCE() to supply this value either directly or using a CELL that contains the GOOGLEFINANCE value.
        *  'yieldpct' does not work for STOCKS and ETF's in GOOGLEFINANCE, so don't supply the third parameter when using that attribute.
    *  Example: (symbol that is not recognized by GOOGLEFINANCE)
        *  ```=CACHEFINANCE("TSE:ZTL", "price", GOOGLEFINANCE("TSE:ZTL", "price"))```


## Using through a trigger.
* The custom function **=CACHEFINANCE()** works well enough, but it is still not 100% (the dreaded `Loading` error)
* Using a trigger ensure that you will **NEVER** have invalid data in your output columns.
  * However, if you provide a stock symbol not recognized by Google AND by any financial website we query - you just won't have any data in that case.
* To use a trigger, you need to create a named range called **CACHEFINANCE**.
  * Go to 'Data' ==> 'Named Ranges' and select job records as the range.  That would be the light green section in the picture below.
  * The named range can be on any sheet and not necessarily on the same sheet as the results.

![Trigger Setup](img/CACHEFINANCE_LEGEND.png)

* Fields in the named range are:
  * **Symbol Range**.  Specify the column range where the stock symbols are located.
  * **Attribute**.  Currently only 'Price', 'Name' and 'yieldpct' are supported.
  * **Output Range**.  Specify the column range where the financial data will be updated.
    * **WARNING** - 1)  There must be the exact same number of cells referenced here as the **Symbol Range**.
    * 2)  This updated by a Trigger function - which does not have the same restrictions as a custom function.  So if you specify a range that overwrites valid data - it will overwrite your valid data.
  * **Google Finance Range**.  This is optional, but recomended.  This column should be the finance data as retrieved by =GOOGLEFINANCE().  Again the number of cells referenced must match the number specified in the **Symbol Range**.
  * **Refresh Minutes**.  When a job is run and data is refreshed, this is the MINIMUM number of minutes to wait before running again.  
  * **Hours**.  The hours of the day when the job can run.  
    * Valid input is 0 to 23.
    * Specific hours can be listed, separated by commas.  e.g.  ```1,9,17,23```
    * Hour ranges be specified with a dash.  e.g.  ```9-17```.
    * ***NOTE**.  Google will interpret 9-17 as a date, so you will need to 'Format' ==> 'Number' ==> 'Plain Text'.
  * **Days of Week**.  Days of the week when the job can run.
    * Valid input is 0,1,2,3,4,5,6  and SUN, MON, TUE, WED, THU, FRI, SAT.
    * Days can be selected using comma separator.  e.g.  ```0-1``` (for Sunday and Monday)
    * A day range can be entered using the dask.  e.g.  ```MON-FRI```.
  * **Trigger ID***.  Just leave BLANK.  This is used by the trigger that starts so it will know which job to take.
  * **NOTE** - Google caps the number of trigger minutes per day, so don't run more often than needed.  For example, how often to Stock Names change?  

* The **Trigger Exists** in the picture above is actually a custom function:  **=CacheFinanceBoot()**
  * This check to make sure that at least one instance of **CacheFinanceTrigger** is set up in the Triggers.
  * If none are found, it will create one (assuming you have the rights - you may need to manually **Run** first).

* **Here is an example of usage - using the CacheFinance Legend above.**
  * Column 'A' is stock symbols entered by you!
  * Column 'B', 'C' and 'D' are updated by the trigger.
  * Column 'F' is entered with:  =GOOGLEFINANCE(A3, "price")
  * Column 'G' is entered with   =GOOGLEFINANCE(A3,"name")

![Trigger Setup](img/ExampleStocks.png)

*  Here are the suggested column titles.
"CacheFinance Legend"	=CacheFinanceBoot()						
"SymbolRange"	"Attribute"	"OutputRange"	"GoogleFinanceRange"	"Refresh Minutes"	"Hours"	"Days of Week"	"Trigger ID"
