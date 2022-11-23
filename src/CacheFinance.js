const GOOGLEFINANCE_PARAM_NOT_USED = "##NotSet##";

//  Named range in sheet with CacheFinance configurations.
const CACHE_LEGEND = "CACHEFINANCE";

/**
 * Add this to App Script Trigger.
 * It requires a named range in your sheet called 'CACHEFINANCE'
 * @param {GoogleAppsScript.Events.TimeDriven} e 
 */
function CacheFinanceTrigger(e) {
    Logger.log("Starting CacheFinanceTrigger()");

    let cacheSettings = new CacheJobSettings();

    //  On first run, it should create a job for every range specified in CACHEFINANCE legend.
    cacheSettings.validateTriggerIDs();
    cacheSettings.createMissingTriggers(60);

    //  Is this job specified in legend.
    /** @type {CacheJob} */
    let jobInfo = cacheSettings.getMyJob(e);

    if (jobInfo === null && (typeof e !== 'undefined')) {
        //  This is a boot job that won't run again.
        cacheSettings.deleteOldTrigger(e.triggerUid);
        return;
    }

    //  Run job to update finance data.
    if (CacheTrigger.runJob(jobInfo)) {
        cacheSettings.deleteOldTrigger(jobInfo.triggerID);
        cacheSettings.createTrigger(jobInfo, jobInfo.getMinutesToNextRun() * 60);
    }

    //  If things go sideways, there are always leftover disabled triggers.
    cacheSettings.cleanupDisabledTriggers();
}

/**
 * Replacement function to GOOGLEFINANCE for stock symbols not recognized by google.
 * @param {string} symbol 
 * @param {string} attribute - ["price", "yieldpct", "name"] 
 * @param {any} googleFinanceValue - Optional.  Use GOOGLEFINANCE() to get value, if '#N/A' will read cache.
 * @returns {any}
 * @customfunction
 */
function CACHEFINANCE(symbol, attribute = "price", googleFinanceValue = GOOGLEFINANCE_PARAM_NOT_USED) {
    Logger.log("CACHEFINANCE:" + symbol + "=" + attribute + ". Google=" + googleFinanceValue);

    if (symbol === '' || attribute === '')
        return '';

    return CacheFinance.getFinanceData(symbol, attribute, googleFinanceValue);
}

/**
 * Add a custom function to your sheet to get the Trigger(s) installed.
 * The customfunction cannot modify the sheet settings for our jobs, so
 * the initial trigger for CacheFinanceTrigger creates the jobs and ends.
 * The actual jobs will run after that.
 * @returns {String[][]}
 * @customfunction
 */
function CacheFinanceBoot() {
    if (CacheJobSettings.bootstrapTrigger())
        return [["Trigger Created!"]];
    else
        return [["Trigger Exists"]];
}

class CacheJobSettings {
    constructor() {
        const sheetNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(CACHE_LEGEND);

        if (sheetNamedRange === null) {
            Logger.log("Named Range CACHEFINANCE not found.");
            Logger.log("Each definition line must define:  'SymbolRange', 'Attribute', 'OutputRange', 'GoogleFinanceRange', 'Refresh Minutes', 'Trigger ID'");
            throw new Error("Named Range CACHEFINANCE not defined.");
        }

        this.jobs = [];
        this.cacheInfo = sheetNamedRange.getValues();

        for (let job of this.cacheInfo) {
            let cacheJob = new CacheJob(job);

            this.jobs.push(cacheJob);
        }
    }

    /**
     * 
     * Adds 'CacheFinanceTrigger' function to triggers.
     */
    static bootstrapTrigger() {
        let missingTrigger = true;
        let validTriggerList = ScriptApp.getProjectTriggers();
        for (let trigger of validTriggerList) {
            // @ts-ignore
            if (trigger.getHandlerFunction().toUpperCase() === 'CACHEFINANCETRIGGER' && !trigger.isDisabled())
                missingTrigger = false;
        }

        if (missingTrigger) {
            Logger.log("Creating BOOTSTRAP Trigger for CacheFinanceTrigger function.")
            ScriptApp
                .newTrigger('CacheFinanceTrigger')
                .timeBased()
                .after(15000)
                .create();

            return true;
        }

        return false;
    }

    cleanupDisabledTriggers() {
        let validTriggerList = ScriptApp.getProjectTriggers();
        for (let trigger of validTriggerList) {
            // @ts-ignore
            if (trigger.getHandlerFunction().toUpperCase() === 'CACHEFINANCETRIGGER' && trigger.isDisabled()) {
                Logger.log("Trigger CLEANUP.  Deleting disabled trigger.");
                ScriptApp.deleteTrigger(trigger);
            }
        }
    }

    validateTriggerIDs() {
        Logger.log("Starting validateTriggerIDs()");

        let validTriggerList = ScriptApp.getProjectTriggers();

        for (let job of this.jobs) {
            if (job.triggerID === "")
                continue;

            let good = false;
            for (let validID of validTriggerList) {
                if (job.triggerID === validID.getUniqueId()) {
                    good = true;
                    break;
                }
            }

            if (!good) {
                Logger.log("Invalid Trigger ID=" + job.triggerID);
                job.triggerID = "";
            }
            else {
                Logger.log("Valid Trigger ID=" + job.triggerID);
            }
        }
    }

    /**
     * 
     * @param {any} triggerID 
     */
    deleteOldTrigger(triggerID) {
        const triggers = ScriptApp.getProjectTriggers();

        let triggerObject = null;
        for (let item of triggers) {
            if (item.getUniqueId() === triggerID)
                triggerObject = item;
        }

        if (triggerObject !== null) {
            Logger.log("DELETING Trigger: " + triggerID);
            ScriptApp.deleteTrigger(triggerObject);
        }
        else {
            Logger.log("Failed to locate trigger to delete: " + triggerID);
        }
    }

    /**
     * 
     * @param {Number} startAfterSeconds 
     */
    createMissingTriggers(startAfterSeconds) {
        for (const job of this.jobs) {
            if (job.triggerID === "" && job.isValidJob()) {
                this.createTrigger(job, startAfterSeconds);
                startAfterSeconds += 60;
            }
        }
    }

    /**
     * 
     * @param {CacheJob} job 
     * @param {Number} startAfterSeconds 
     */
    createTrigger(job, startAfterSeconds) {
        if (job.timeout) {
            startAfterSeconds = 15;
        }

        let newTriggerID = ScriptApp
            .newTrigger('CacheFinanceTrigger')
            .timeBased()
            .after(startAfterSeconds * 1000)
            .create();
        job.triggerID = newTriggerID.getUniqueId();

        if (job.timeout) {
            let savedPartialResults = JSON.stringify(job.jobData);
            let key = CacheFinance.makeCacheKey("RESUMEJOB", job.triggerID);
            const shortCache = CacheService.getScriptCache();
            shortCache.put(key, savedPartialResults, 300);
            Logger.log("Saving PARTIAL RESULTS.  Key=" + key + ". Items=" + job.jobData.length);
        }

        this.updateFinanceCacheLegend();
        Logger.log("New Trigger Created. ID=" + job.triggerID + ". Attrib=" + job.attribute + ". Start in " + startAfterSeconds + " seconds.");
    }

    updateFinanceCacheLegend() {
        let legend = [];

        for (let job of this.jobs) {
            legend.push(job.toArray());
        }

        Logger.log("Updating LEGEND." + legend);

        const sheetNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(CACHE_LEGEND);
        sheetNamedRange.setValues(legend);
    }

    /**
     * 
     * @param {GoogleAppsScript.Events.TimeDriven} e 
     * @returns {CacheJob}
     */
    getMyJob(e) {
        if (typeof e === 'undefined') {
            Logger.log("Trigger ID unknown.  Function not started as a TimeDriven event");
            return null;
        }

        let myJob = null;
        const runningScriptID = e.triggerUid;

        for (const job of this.jobs) {
            if (job.triggerID === runningScriptID) {
                Logger.log("Found My Job ID=" + job.triggerID);
                myJob = job;
                break;
            }
        }

        if (myJob === null) {
            Logger.log("This JOB " + runningScriptID + " not found in legend.");
        }
        else {
            //  Is this a continuation job?
            let key = CacheFinance.makeCacheKey("RESUMEJOB", myJob.triggerID);
            const shortCache = CacheService.getScriptCache();
            let resumeData = shortCache.get(key);
            if (resumeData !== null) {
                myJob.wasRestarted = true;
                myJob.jobData = JSON.parse(resumeData);
                Logger.log("Resuming Job: " + key + ". Items added=" + myJob.jobData.length);
            }
            else {
                Logger.log("New Job.  CacheKey=" + key);
            }
        }

        return myJob;
    }
}

class CacheJob {
    /**
     * 
     * @param {any[]} jobParmameters 
     */
    constructor(jobParmameters) {
        this.DAYNAMES = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];

        const CACHE_SETTINGS_SYMBOL = 0;
        const CACHE_SETTINGS_ATTRIBUTE = 1;
        const CACHE_SETTINGS_OUTPUT = 2;
        const CACHE_SETTINGS_DEFAULT = 3;
        const CACHE_SETTINGS_REFRESH = 4;
        const CACHE_SETTINGS_HOURS = 5;
        const CACHE_SETTINGS_DAYS = 6;
        const CACHE_SETTINGS_ID = 7;

        this.symbolRange = jobParmameters[CACHE_SETTINGS_SYMBOL];
        this.attribute = jobParmameters[CACHE_SETTINGS_ATTRIBUTE];
        this.outputRange = jobParmameters[CACHE_SETTINGS_OUTPUT];
        this.defaultRange = jobParmameters[CACHE_SETTINGS_DEFAULT];
        this.refreshMinutes = jobParmameters[CACHE_SETTINGS_REFRESH];
        this.hours = jobParmameters[CACHE_SETTINGS_HOURS].toString();
        this.days = jobParmameters[CACHE_SETTINGS_DAYS].toString();
        this.triggerID = jobParmameters[CACHE_SETTINGS_ID];
        this.dayNumbers = [];
        this.hourNumbers = [];
        this.jobData = null;
        this.wasRestarted = false;

        Logger.log("Job Settings:  symbols=" + this.symbolRange + ". attribute=" + this.attribute + ". Out=" + this.outputRange + ". In=" + this.defaultRange + ". Minutes=" + this.refreshMinutes + ". Hours=" + this.hours + ". Days=" + this.days);

        if (typeof this.hours === 'string')
            this.hours = this.hours.trim().toUpperCase();

        if (typeof this.days === 'string')
            this.days = this.days.trim().toUpperCase();

        this.extractRunDaysOfWeek();
        this.extractHoursOfDay();

        Logger.log("CACHEFINANCE JOB. Symbol=" + this.symbolRange + ". Attrib=" + this.attribute + ". Output=" + this.outputRange + ". Refresh=" + this.refreshMinutes);
    }

    isValidJob() {
        return this.symbolRange !== '' && this.outputRange !== '' && this.attribute !== '' && this.dayNumbers.length > 0 && this.hourNumbers.length > 0;
    }

    /**
     * 
     * @param {any[][]} data 
     */
    save(data) {
        this.jobData = data;
    }

    timedOut(value) {
        this.timeout = value;
    }

    isRestarting() {
        return this.wasRestarted;
    }

    extractRunDaysOfWeek() {
        this.dayNumbers = [];

        if (this.days === '') {
            this.days = '*';
        }

        for (let i = 0; i < 7; i++)
            this.dayNumbers[i] = false;

        if (this.days === '*') {
            this.days = this.DAYNAMES.join(",");
        }

        if (this.extractListDaysOfWeek()) {
            return;
        }

        if (this.extractRangeDaysOfWeek()) {
            return;
        }

        let dayNameIndex = this.DAYNAMES.indexOf(this.days);
        if (dayNameIndex !== -1) {
            this.dayNumbers[dayNameIndex] = true;
            return;
        }

        let singleItem = parseInt(this.days);
        if (!isNaN(singleItem)) {
            this.dayNumbers[singleItem] = true;
            return;
        }

        Logger.log("Job does not contain any days of the week to run. " + this.days);
    }

    extractListDaysOfWeek() {
        if (this.days.indexOf(",") === -1) {
            return false;
        }

        let listValues = this.days.split(",");
        listValues = listValues.map(p => p.trim());

        for (let day of listValues) {
            let dayNameIndex = this.DAYNAMES.indexOf(day);
            if (dayNameIndex == -1) {
                dayNameIndex = parseInt(day);
                if (isNaN(dayNameIndex))
                    dayNameIndex = -1;

                if (dayNameIndex < 0 || dayNameIndex > 6)
                    dayNameIndex = -1;
            }

            if (dayNameIndex !== -1) {
                this.dayNumbers[dayNameIndex] = true;
            }
        }

        return true;
    }

    extractRangeDaysOfWeek() {
        if (this.days.indexOf("-") === -1) {
            return false;
        }

        let rangeValues = this.days.split("-");
        rangeValues = rangeValues.map(p => p.trim());

        let startDayNameIndex = this.DAYNAMES.indexOf(rangeValues[0]);
        let endDayNameIndex = this.DAYNAMES.indexOf(rangeValues[1]);
        if (startDayNameIndex == -1 || endDayNameIndex == -1) {
            startDayNameIndex = parseInt(rangeValues[0]);
            endDayNameIndex = parseInt(rangeValues[1]);

            if (isNaN(startDayNameIndex) || isNaN(endDayNameIndex)) {
                return;
            }
        }

        let count = 0;
        for (let i = startDayNameIndex; count < 7; count++) {
            this.dayNumbers[i] = true;

            if (i == endDayNameIndex)
                break;

            i++;
            if (i > 6)
                i = 0;
        }

        return true;
    }

    extractHoursOfDay() {
        this.hourNumbers = [];

        if (this.hours === '') {
            this.hours = '*';
        }

        for (let i = 0; i < 24; i++)
            this.hourNumbers[i] = (this.hours === '*') ? true : false;

        if (this.hours === '*')
            return;

        if (this.extractListHoursOfDay())
            return;


        if (this.extractRangeHoursOfDay())
            return;

        let singleItem = parseInt(this.hours);
        if (!isNaN(singleItem)) {
            this.hourNumbers[singleItem] = true;
            return;
        }

        Logger.log("This job does not contain any valid hours to run. " + this.hours);
    }

    extractListHoursOfDay() {
        if (this.hours.indexOf(",") === -1) {
            return false;
        }

        let listValues = this.hours.split(",");
        listValues = listValues.map(p => p.trim());

        for (let hr of listValues) {
            let hourIndex = parseInt(hr);
            if (isNaN(hourIndex))
                hourIndex = -1;

            if (hourIndex < 0 || hourIndex > 23)
                hourIndex = -1;


            if (hourIndex !== -1) {
                this.hourNumbers[hourIndex] = true;
            }
        }

        return true;
    }

    extractRangeHoursOfDay() {
        if (this.hours.indexOf("-") === -1) {
            return false;
        }

        let rangeValues = this.hours.split("-");
        rangeValues = rangeValues.map(p => p.trim());

        let startHourIndex = parseInt(rangeValues[0]);
        let endHourIndex = parseInt(rangeValues[1]);

        if (isNaN(startHourIndex) || isNaN(endHourIndex)) {
            return;
        }

        if (startHourIndex < 0 || startHourIndex > 23 ||
            endHourIndex < 0 || endHourIndex > 23) {
            return;
        }

        let count = 0;
        for (let i = startHourIndex; count < 24; count++) {
            this.hourNumbers[i] = true;

            if (i == endHourIndex)
                break;

            i++;
            if (i > 23)
                i = 0;
        }

        return true;
    }

    getMinutesToNextRun() {
        let minutes = this.refreshMinutes < 1 ? 1 : this.refreshMinutes;
        // Get current date
        const startDateTime = new Date();
        let daysSinceStart = 0;
        let date = new Date();
        date.setMinutes(date.getMinutes() + minutes);

        //  The next run is within our window of opportunity.
        while (!this.canRunJob(date) && daysSinceStart <= 7) {
            if (!this.canRunJobForDayOfWeek(date) ||
                (this.canRunJobForDayOfWeek(date) && !this.canRunJobLaterToday(date))) {
                date.setDate(date.getDate() + 1);
                date.setHours(0);
                date.setMinutes(0)
            }
            else {
                date.setHours(this.getNextHourToRun(date));
                date.setMinutes(0);
            }

            daysSinceStart = (date.getTime() - startDateTime.getTime()) / (1000 * 3600 * 24);
        }

        if (daysSinceStart > 7) {
            throw new Error("Error finding next TRIGGER time for job.");
        }

        minutes = (date.getTime() - startDateTime.getTime()) / 60000;

        return minutes;
    }

    /**
     * 
     * @param {Date} startDate 
     * @returns {boolean}
     */
    canRunJob(startDate = new Date()) {
        return this.canRunJobForDayOfWeek(startDate) && this.canRunJobForTime(startDate);
    }

    /**
     * 
     * @param {Date} date 
     * @returns {Boolean}
     */
    canRunJobForDayOfWeek(date) {
        let dow = date.getDay();
        let status = this.dayNumbers[dow];

        return status;
    }

    /**
     * 
     * @param {Date} date 
     * @returns {Boolean}
     */
    canRunJobForTime(date) {
        let hour = date.getHours();
        let status = this.hourNumbers[hour];

        return status;
    }

    /**
     * 
     * @param {Date} date 
     * @returns {Boolean}
     */
    canRunJobLaterToday(date) {
        return this.getNextHourToRun(date) !== -1;
    }

    /**
     * 
     * @param {Date} date 
     * @returns {Number}
     */
    getNextHourToRun(date) {
        let hr = date.getHours();
        for (let i = hr; i < 24; i++) {
            if (this.hourNumbers[i])
                return i;
        }
        return -1;
    }

    /**
     * 
     * @returns {any[]}
     */
    toArray() {
        let row = [];
        row.push(this.symbolRange);
        row.push(this.attribute);
        row.push(this.outputRange);
        row.push(this.defaultRange);
        row.push(this.refreshMinutes);
        row.push(this.hours);
        row.push(this.days);
        row.push(this.triggerID);

        return row;
    }
}

class CacheTrigger {
    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {Boolean} - true - DELETE JOB and RE-CREATE.
     */
    static runJob(jobSettings) {
        if (jobSettings === null) {
            Logger.log("Job settings not found.");
            return false;
        }

        if (!jobSettings.canRunJob()) {
            Logger.log("Not time to run JOB.")
            return true;
        }

        Logger.log("Starting JOB.  Updating FINANCE info.");

        CacheTrigger.getJobData(jobSettings);
        CacheTrigger.writeResults(jobSettings);

        return true;
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     */
    static getJobData(jobSettings) {
        const MAX_RUN_SECONDS = 300;
        let data = [];
        const symbolsNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.symbolRange);
        let googleValuesNamedRange = null;
        if (jobSettings.defaultRange !== "") {
            googleValuesNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.defaultRange);
        }

        Logger.log("CacheTrigger: Symbols:" + jobSettings.symbolRange + ". Attribute:" + jobSettings.attribute + ". GoogleRange: " + jobSettings.defaultRange);

        if (symbolsNamedRange === null || (jobSettings.defaultRange !== "" && googleValuesNamedRange === null)) {
            Logger.log("Failed to read data from range.");
            return;
        }

        const symbols = symbolsNamedRange.getValues();
        let defaultData = [];
        if (googleValuesNamedRange !== null) {
            defaultData = googleValuesNamedRange.getValues();
        }

        const attribute = jobSettings.attribute;

        if (googleValuesNamedRange != null && symbols.length !== defaultData.length) {
            Logger.log("Symbol Ranges and Google Values Ranges must be the same: " + jobSettings.symbolRange + ". Len=" + symbols.length + ".  GoogleValues: " + jobSettings.defaultRange + ". Len=" + defaultData.length);
            return;
        }

        jobSettings.timedOut(false);

        let startingSymbol = 0;
        if (jobSettings.isRestarting()) {
            data = jobSettings.jobData;
            startingSymbol = data.length;
        }

        const start = new Date().getTime();

        for (let i = startingSymbol; i < symbols.length; i++) {
            let elapsed = (new Date().getTime() - start) / 1000;
            if (elapsed > MAX_RUN_SECONDS) {
                Logger.log("Max. Job Time reached.");
                jobSettings.timedOut(true);
                break;
            }

            let value;
            if (googleValuesNamedRange === null) {
                Logger.log("symbol=" + symbols[i][0]);
                value = CACHEFINANCE(symbols[i][0], attribute);
            }
            else {
                Logger.log("symbol=" + symbols[i][0] + ". Default=" + defaultData[i][0]);
                value = CACHEFINANCE(symbols[i][0], attribute, defaultData[i][0]);
            }

            data.push([value]);
        }

        jobSettings.save(data);
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns 
     */
    static writeResults(jobSettings) {
        if (jobSettings.jobData === null || jobSettings.timeout)
            return false;

        Logger.log("writeCacheResults:  START.  Data Len=" + jobSettings.jobData.length);

        const outputNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.outputRange);

        if (outputNamedRange === null) {
            return false;
        }

        try {
            outputNamedRange.setValues(jobSettings.jobData);
        }
        catch (ex) {
            Logger.log("Updating output range FAILED.  " + ex.toString());
            return false;
        }

        Logger.log("writeCacheResults:  END: " + jobSettings.outputRange);
        return true;
    }
}

class CacheFinance {
    /**
     * Replacement function to GOOGLEFINANCE for stock symbols not recognized by google.
     * @param {string} symbol 
     * @param {string} attribute - ["price", "yieldpct", "name"] 
     * @param {any} googleFinanceValue - Optional.  Use GOOGLEFINANCE() to get value, if '#N/A' will read cache.
     * @returns {any}
     */
    static getFinanceData(symbol, attribute, googleFinanceValue) {
        attribute = attribute.toUpperCase().trim();
        symbol = symbol.toUpperCase();
        const cacheKey = CacheFinance.makeCacheKey(symbol, attribute);

        //  This time GOOGLEFINANCE worked!!!
        if (googleFinanceValue !== GOOGLEFINANCE_PARAM_NOT_USED && googleFinanceValue !== "#N/A" && googleFinanceValue !== '#ERROR!') {
            //  We cache here longer because we would normally be getting data from Google.
            //  If GoogleFinance is failing, we need the data to be held longer since it
            //  it is getting from cache as an emergency backup.
            CacheFinance.saveFinanceValueToCache(cacheKey, googleFinanceValue, 21600);
            return googleFinanceValue;
        }

        //  In the case where GOOGLE never gives a value (or GoogleFinance is never used),
        //  we don't want to pull from long cache (at this point).
        const useShortCacheOnly = googleFinanceValue === GOOGLEFINANCE_PARAM_NOT_USED || googleFinanceValue !== "#N/A";

        //  GOOGLEFINANCE has failed OR was not used.  Is it in the cache?
        let data = CacheFinance.getFinanceValueFromCache(cacheKey, useShortCacheOnly);

        if (data !== null)
            return data;

        //  Last resort... try other sites.
        let stockAttributes = ThirdPartyFinance.get(symbol, attribute);

        //  Failed third party lookup, try using long term cache.
        if (!stockAttributes.isAttributeSet(attribute)) {
            let cachedStockAttribute = CacheFinance.getFinanceValueFromCache(cacheKey, false);
            if (cachedStockAttribute !== null)
                stockAttributes = cachedStockAttribute;
        }
        else {
            //  If we are mostly getting this finance item from a third party, we set the timeout
            //  a little shorter since we don't want to return extremely old data.
            CacheFinance.saveAllFinanceValuesToCache(symbol, stockAttributes);
        }

        return stockAttributes.getValue(attribute);
    }

    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {String}
     */
    static makeCacheKey(symbol, attribute) {
        return attribute + "|" + symbol;
    }

    /**
     * 
     * @param {String} cacheKey 
     * @param {Boolean} useShortCacheOnly
     * @returns {any}
     */
    static getFinanceValueFromCache(cacheKey, useShortCacheOnly) {
        const shortCache = CacheService.getScriptCache();
        const longCache = new ScriptSettings();

        let data = shortCache.get(cacheKey);

        //  Set to null while testing.  Remove when all is working.
        // data = null;

        if (data !== null && data !== "#ERROR!") {
            Logger.log("Found in Short CACHE: " + cacheKey + ". Value=" + data);
            let parsedData = JSON.parse(data);
            if (!(typeof parsedData === 'string' && parsedData === "#ERROR!"))
                return parsedData;
        }

        if (useShortCacheOnly)
            return null;

        data = longCache.get(cacheKey);
        if (data !== null && data !== "#ERROR!") {
            Logger.log("Long Term Cache.  Key=" + cacheKey + ". Value=" + data);
            //  Long cache saves and returns same data type -so no conversion needed.
            return data;
        }

        return null;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {StockAttributes} stockAttributes 
     */
    static saveAllFinanceValuesToCache(symbol, stockAttributes) {
        if (stockAttributes === null)
            return;
        if (stockAttributes.stockName !== null)
            CacheFinance.saveFinanceValueToCache(CacheFinance.makeCacheKey(symbol, "NAME"), stockAttributes.stockName, 1200);
        if (stockAttributes.stockPrice !== null)
            CacheFinance.saveFinanceValueToCache(CacheFinance.makeCacheKey(symbol, "PRICE"), stockAttributes.stockPrice, 1200);
        if (stockAttributes.yieldPct !== null)
            CacheFinance.saveFinanceValueToCache(CacheFinance.makeCacheKey(symbol, "YIELDPCT"), stockAttributes.yieldPct, 1200);
    }

    /**
     * 
     * @param {String} key 
     * @param {any} financialData 
     * @param {Number} shortCacheSeconds 
     */
    static saveFinanceValueToCache(key, financialData, shortCacheSeconds = 1200) {
        const shortCache = CacheService.getScriptCache();
        const longCache = new ScriptSettings();
        let start = new Date().getTime();

        Logger.log("GoogleFinance VALUE set to Long/Short Cache. Key=" + key + ".  Value=" + financialData);

        const currentShortCacheValue = shortCache.get(key);
        if (currentShortCacheValue !== null && JSON.parse(currentShortCacheValue) === financialData) {
            Logger.log("GoogleFinance VALUE.  No Change in SHORT Cache. ms=" + (new Date().getTime() - start));
            return;
        }

        if (currentShortCacheValue !== null)
            Logger.log("Short Cache Changed.  Old=" + JSON.parse(currentShortCacheValue) + " . New=" + financialData);

        start = new Date().getTime();

        //  If we normally get the price from Google, we want to cache for a longer
        //  time because the only time we need a price for this particular stock
        //  is when GOOGLEFINANCE fails.
        shortCache.put(key, JSON.stringify(financialData), shortCacheSeconds);

        Logger.log("Time to put in short cache: " + (new Date().getTime() - start));
        start = new Date().getTime();

        //  For emergency cases when GOOGLEFINANCE is down long term...
        longCache.put(key, financialData, 7);
        Logger.log("Time to put in long cache: " + (new Date().getTime() - start));
    }
}

class ThirdPartyFinance {
    /**
     * 
     * @param {String} symbol 
     * @param {String} attribute 
     * @returns {StockAttributes}
     */
    static get(symbol, attribute) {
        let data = null;

        switch (attribute) {
            case "PRICE":
                data = ThirdPartyFinance.getStockPrice(symbol);
                break;

            case "YIELDPCT":
                data = ThirdPartyFinance.getStockDividendYield(symbol);
                break;

            case "NAME":
                data = ThirdPartyFinance.getName(symbol);
                break;

            default:
                Logger.log("Invalid FINANCE attribute: " + attribute);
                throw new Error("Invalid attribute:" + attribute);
        }

        if (data.stockPrice !== null)
            data.stockPrice = Math.round(data.stockPrice * 100) / 100;

        if (data.yieldPct !== null)
            data.yieldPct = Math.round(data.yieldPct * 10000) / 10000;

        return data;
    }

    /**
     * 
     * @param {string} symbol 
     * @returns {StockAttributes}
     */
    static getStockPrice(symbol) {
        /** @type {StockAttributes} */
        let data = GlobeAndMail.getInfo(symbol);

        if (data.stockPrice === null)
            data = TdMarketResearch.getInfo(symbol, "ETF");

        if (data.stockPrice === null)
            data = TdMarketResearch.getInfo(symbol, "STOCK");

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getStockDividendYield(symbol) {

        let data = GlobeAndMail.getInfo(symbol);

        if (data.yieldPct === null)
            data = TdMarketResearch.getInfo(symbol, "ETF");

        if (data.yieldPct === null)
            data = TdMarketResearch.getInfo(symbol, "STOCK");

        if (data.yieldPct === null)
            data = YahooFinance.getInfo(symbol);

        return data;
    }

    static getName(symbol) {
        /** @type {StockAttributes} */
        let data = GlobeAndMail.getInfo(symbol);

        if (data.stockName === null)
            data = TdMarketResearch.getInfo(symbol, "ETF");

        if (data.stockName === null)
            data = TdMarketResearch.getInfo(symbol, "STOCK");

        return data;
    }

}

function testGlobe() {
    let data = GlobeAndMail.getInfo("TSE:FTN-A");
    Logger.log("Globe: FTN-A=" + data);

    data = GlobeAndMail.getInfo("TSE:HBF.B");
    Logger.log("Globe HBF.B=" + data);

    data = GlobeAndMail.getInfo("TSE:MEG");
    Logger.log("Globe MEG=" + data);
}

function testYahooDividend() {
    let div = YahooFinance.getInfo("NYSEARCA:SHYG");
    Logger.log("Dividend NYSEARCA:SHYG=" + div);
}

function testgetTDmarketResearchName() {
    let coName = TdMarketResearch.getInfo("ZTL");

    Logger.log("Name of ZtL: " + coName);
}

function testTD() {
    TdMarketResearch.getInfo("TSE:DFN-A", "STOCK");
}


function testgetTDmarketResearchPrice() {
    TdMarketResearch.getInfo("ZTL", "ETF");
}

class TdMarketResearch {
    /**
     * 
     * @param {String} symbol 
     * @param {String} type 
     * @returns {StockAttributes}
     */
    static getInfo(symbol, type = "ETF") {
        let data = new StockAttributes();

        let URL;
        if (type == "ETF")
            URL = "https://marketsandresearch.td.com/tdwca/Public/ETFsProfile/Summary/ca/" + TdMarketResearch.getTicker(symbol);
        else
            URL = "https://marketsandresearch.td.com/tdwca/Public/Stocks/Overview/ca/" + TdMarketResearch.getTicker(symbol);

        let html;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return data;
        }
        Logger.log("getStockDividendYield:  " + symbol);
        Logger.log("URL = " + URL);

        //  Get the dividend yield.
        let parts = html.match(/.Dividend Yield\<\/th\>\<td class\=\"last\"\>(\d*\.?\d*)\%/);
        if (parts === null) {
            parts = html.match(/.Dividend Yield\<\/div\>.*?cell-container contains\"\>(\d*\.?\d*)\%/);
        }
        if (parts !== null && parts.length == 2) {
            let tempPct = parts[1];

            let parsedValue = parseFloat(tempPct) / 100;

            if (!isNaN(parsedValue)) {
                data.yieldPct = parsedValue;
            }
        }

        //  Get the name.
        parts = html.match(/.\<span class=\"issueName\"\>(.*?)\<\//);
        if (parts !== null && parts.length == 2) {
            data.stockName = parts[1];
        }

        //  Get the price.
        parts = html.match(/.LAST PRICE\<\/span\>\<div\>\<span\>(\d*\.?\d*)\</);
        if (parts !== null && parts.length == 2) {

            let parsedValue = parseFloat(parts[1]);
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
            let parts = symbol.split(":");
            symbol = parts[1];
        }

        const dash = symbol.indexOf("-");
        if (dash >= 0) {
            symbol = symbol.substr(0, dash) + ".PR." + symbol.substr(dash + 1);
        }

        return symbol;
    }
}


class YahooFinance {
    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        let data = new StockAttributes();

        let URL = "https://finance.yahoo.com/quote/" + YahooFinance.getTicker(symbol);

        const html = UrlFetchApp.fetch(URL).getContentText();
        Logger.log("getStockDividendYield:  " + symbol);
        Logger.log("URL = " + URL);

        let dividendPercent = html.match(/.TD_YIELD-value\"\>(\d*\.?\d*)\%/);

        if (dividendPercent !== null && dividendPercent.length == 2) {
            const tempPct = dividendPercent[1];
            Logger.log("PERCENT=" + tempPct);

            data.yieldPct = parseFloat(tempPct) / 100;

            if (isNaN(data.yieldPct)) {
                data.yieldPct = null;
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
        let colon = symbol.indexOf(":");

        if (colon >= 0) {
            const symbolParts = symbol.split(":");

            modifiedSymbol = symbolParts[1];
            if (symbolParts[0] === "TSE")
                modifiedSymbol = symbolParts[1] + ".TO";

        }
        return modifiedSymbol;
    }
}

class GlobeAndMail {
    /**
     * Only gets dividend yield.
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getInfo(symbol) {
        let data = new StockAttributes();
        let URL = "https://www.theglobeandmail.com/investing/markets/stocks/" + GlobeAndMail.getTicker(symbol);

        Logger.log("getStockDividendYield:  " + symbol);
        Logger.log("URL = " + URL);

        let html;
        try {
            html = UrlFetchApp.fetch(URL).getContentText();
        }
        catch (ex) {
            return data;
        }

        //  Get the dividend yield.
        let parts = html.match(/.name=\"dividendYieldTrailing\".*?value=\"(\d*\.?\d*)\%/);

        if (parts === null)
            parts = html.match(/.name=\\\"dividendYieldTrailing\\\".*?value=\\\"(\d*\.?\d*)\%/);

        if (parts !== null && parts.length == 2) {
            let tempPct = parts[1];

            let parsedValue = parseFloat(tempPct) / 100;

            if (!isNaN(parsedValue)) {
                data.yieldPct = parsedValue;
            }
        }

        //  Get the name.
        parts = html.match(/.\"symbolName\":\"(.*?)\"/);
        if (parts !== null && parts.length == 2) {
            data.stockName = parts[1];
        }

        //  Get the price.
        parts = html.match(/.\"lastPrice\":\"(\d*\.?\d*)\"/);
        if (parts !== null && parts.length == 2) {

            let parsedValue = parseFloat(parts[1]);
            if (!isNaN(parsedValue)) {
                data.stockPrice = parsedValue;
            }
        }


        return data;
    }

    //  
    /**
     * Clean up ticker symbol for use in Globe and Mail lookups.
     * @param {String} symbol 
     * @returns {String}
     */
    static getTicker(symbol) {
        let colon = symbol.indexOf(":");

        if (colon >= 0) {
            let parts = symbol.split(":");

            switch (parts[0].toUpperCase()) {
                case "TSE":
                    symbol = parts[1];
                    if (parts[1].indexOf(".") !== -1) {
                        symbol = parts[1].replace(".", "-");
                    }
                    else if (parts[1].indexOf("-") != -1) {
                        const prefShare = parts[1].split("-");
                        symbol = prefShare[0] + "-PR-" + prefShare[1];
                    }
                    symbol = symbol + "-T";
                    break;
                case "NYSEARCA":
                    symbol = parts[1] + "-A";
                    break;
                case "NASDAQ":
                    symbol = parts[1] + "-Q";
                    break;
            }
        }

        return symbol;
    }
}

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
                Logger.log("Invalid FINANCE attribute: " + attribute);
                throw new Error("Invalid attribute:" + attribute);
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
                Logger.log("Invalid FINANCE attribute: " + attribute);
                throw new Error("Invalid attribute:" + attribute);
        }
    }
}




