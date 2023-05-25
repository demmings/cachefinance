const GOOGLEFINANCE_PARAM_NOT_USED = "##NotSet##";

/**
 * Replacement function to GOOGLEFINANCE for stock symbols not recognized by google.
 * @param {string} symbol 
 * @param {string} attribute - ["price", "yieldpct", "name"] - "TEST" - returns test results from 3rd party sites.
 * @param {any} googleFinanceValue - Optional.  Use GOOGLEFINANCE() to get value, if '#N/A' will read cache.
 * @returns {any}
 * @customfunction
 */
function CACHEFINANCE(symbol, attribute = "price", googleFinanceValue = GOOGLEFINANCE_PARAM_NOT_USED) {
    Logger.log(`CACHEFINANCE:${symbol}=${attribute}. Google=${googleFinanceValue}`);

    if (attribute.toUpperCase() === "TEST") {
        return cacheFinanceTest();
    }

    if (symbol === '' || attribute === '') {
        return '';
    }

    return CacheFinance.getFinanceData(symbol, attribute, googleFinanceValue);
}


/**
 * @classdesc GOOGLEFINANCE helper function.  Returns default value (if available) and set this value to cache OR
 * reads from short term cache (<21600s) and returns value OR
 * reads from 3rd party screen scrapping OR
 * reads from long term cache
 */
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
        const data = CacheFinance.getFinanceValueFromCache(cacheKey, useShortCacheOnly);

        if (data !== null)
            return data;

        //  Last resort... try other sites.
        let stockAttributes = ThirdPartyFinance.get(symbol, attribute);

        //  Failed third party lookup, try using long term cache.
        if (!stockAttributes.isAttributeSet(attribute)) {
            const cachedStockAttribute = CacheFinance.getFinanceValueFromCache(cacheKey, false);
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
        return `${attribute}|${symbol}`;
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
            Logger.log(`Found in Short CACHE: ${cacheKey}. Value=${data}`);
            const parsedData = JSON.parse(data);
            if (!(typeof parsedData === 'string' && parsedData === "#ERROR!"))
                return parsedData;
        }

        if (useShortCacheOnly)
            return null;

        data = longCache.get(cacheKey);
        if (data !== null && data !== "#ERROR!") {
            Logger.log(`Long Term Cache.  Key=${cacheKey}. Value=${data}`);
            //  Long cache saves and returns same data type -so no conversion needed.
            return data;
        }

        return null;
    }

    /**
     * 
     * @param {String} symbol 
     * @param {StockAttributes} stockAttributes 
     * @returns {void}
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
     * @returns {void}
     */
    static saveFinanceValueToCache(key, financialData, shortCacheSeconds = 1200) {
        const shortCache = CacheService.getScriptCache();
        const longCache = new ScriptSettings();
        let start = new Date().getTime();

        const currentShortCacheValue = shortCache.get(key);
        if (currentShortCacheValue !== null && JSON.parse(currentShortCacheValue) === financialData) {
            Logger.log(`GoogleFinance VALUE.  No Change in SHORT Cache. ms=${new Date().getTime() - start}`);
            return;
        }

        if (currentShortCacheValue !== null)
            Logger.log(`Short Cache Changed.  Old=${JSON.parse(currentShortCacheValue)} . New=${financialData}`);
   
        //  If we normally get the price from Google, we want to cache for a longer
        //  time because the only time we need a price for this particular stock
        //  is when GOOGLEFINANCE fails.
        start = new Date().getTime();
        shortCache.put(key, JSON.stringify(financialData), shortCacheSeconds);
        const shortMs = new Date().getTime() - start;
       
        //  For emergency cases when GOOGLEFINANCE is down long term...
        start = new Date().getTime();
        longCache.put(key, financialData, 7);
        const longMs = new Date().getTime() - start;

        Logger.log(`SET GoogleFinance VALUE Long/Short Cache. Key=${key}.  Value=${financialData}. Short ms=${shortMs}. Long ms=${longMs}`);
    }
}






//  Named range in sheet with CacheFinance configurations.
const CACHE_LEGEND = "CACHEFINANCE";

/**
 * Add this to App Script Trigger.
 * It requires a named range in your sheet called 'CACHEFINANCE'
 * @param {GoogleAppsScript.Events.TimeDriven} e 
 */
function CacheFinanceTrigger(e) {                           //  skipcq: JS-0128
    Logger.log("Starting CacheFinanceTrigger()");

    const cacheSettings = new CacheJobSettings();

    //  The trigger ID for THIS job is already disabled.  Send a signal to other
    //  running jobs that THIS trigger is still going.
    CacheJobSettings.signalTriggerRunState(e, true);

    //  Is this job specified in legend.
    /** @type {CacheJob} */
    const jobInfo = cacheSettings.getMyJob(e);

    if (jobInfo === null && (typeof e !== 'undefined')) {
        //  This is a boot job that won't run again.
        cacheSettings.firstRun(e.triggerUid);
        return;
    }

    //  Run job to update finance data.
    CacheTrigger.runJob(jobInfo);

    //  Create new job and ensure all existing jobs are valid.
    cacheSettings.afterRun(jobInfo);

    //  Signal outside trigger that this ID is not running.
    CacheJobSettings.signalTriggerRunState(e, false);
}

/**
 * Add a custom function to your sheet to get the Trigger(s) installed.
 * The customfunction cannot modify the sheet settings for our jobs, so
 * the initial trigger for CacheFinanceTrigger creates the jobs and ends.
 * The actual jobs will run after that.
 * @returns {String[][]}
 * @customfunction
 */
function CacheFinanceBoot() {                       //  skipcq: JS-0128
    if (CacheJobSettings.bootstrapTrigger())
        return [["Trigger Created!"]];
    else
        return [["Trigger Exists"]];
}

/**
 * Manage TRIGGERS that run and update stock/etf data from the web.
 */
class CacheJobSettings {
    constructor() {
        this.load(null);
    }

    /**
     * 
     * @param {CacheJob} currentJob 
     */
    load(currentJob) {
        const sheetNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(CACHE_LEGEND);

        if (sheetNamedRange === null) {
            Logger.log("Named Range CACHEFINANCE not found.");
            Logger.log("Each definition line must define:  'SymbolRange', 'Attribute', 'OutputRange', 'GoogleFinanceRange', 'Refresh Minutes', 'Trigger ID'");
            throw new Error("Named Range CACHEFINANCE not defined.");
        }

        /** @type {CacheJob[]} */
        this.jobs = [];

        this.cacheInfo = sheetNamedRange.getValues();

        for (const job of this.cacheInfo) {
            const cacheJob = new CacheJob(job);

            if (currentJob !== null && cacheJob.triggerID === currentJob.triggerID) {
                this.jobs.push(currentJob);
            }
            else {
                this.jobs.push(cacheJob);
            }
        }
    }

    /**
     * 
     * Adds 'CacheFinanceTrigger' function to triggers.
     * @returns {Boolean} - true = created
     */
    static bootstrapTrigger() {
        let missingTrigger = true;
        const validTriggerList = ScriptApp.getProjectTriggers();
        for (const trigger of validTriggerList) {
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

    /**
     * 
     * @param {String} triggerUid 
     * @returns {void}
     */
    firstRun(triggerUid) {
        //  On first run, it should create a job for every range specified in CACHEFINANCE legend.
        this.validateTriggerIDs();
        this.createMissingTriggers(true);

        //  This is the BOOT trigger ID.
        CacheJobSettings.deleteOldTrigger(triggerUid);
    }

    /**
     * 
     * @param {CacheJob} jobInfo 
     * @returns {void}
     */
    afterRun(jobInfo) {
        //  Reload job table in case another trigger recently updated.
        this.load(jobInfo);

        CacheJobSettings.deleteOldTrigger(jobInfo.triggerID);       //  Delete myself
        jobInfo.triggerID = "";                         
        CacheJobSettings.cleanupDisabledTriggers();                 //  Delete triggers that ran, but not cleaned up.    
        this.validateTriggerIDs();
        this.createMissingTriggers(false);
    }

    /**
     * Find trigger that are disabled and not associated with anything running.
     * @returns {void}
     */
    static cleanupDisabledTriggers() {
        const validTriggerList = ScriptApp.getProjectTriggers();
        for (const trigger of validTriggerList) {
            // @ts-ignore
            if (trigger.getHandlerFunction().toUpperCase() === 'CACHEFINANCETRIGGER' && trigger.isDisabled() &&
                !CacheJobSettings.isDisabledTriggerStillRunning(trigger.getUniqueId())) {
                Logger.log("Trigger CLEANUP.  Deleting disabled trigger.");
                ScriptApp.deleteTrigger(trigger);
            }
        }
    }

    /**
     * 
     * @param {GoogleAppsScript.Events.TimeDriven} e 
     * @param {Boolean} alive 
     * @returns {void}
     */
    static signalTriggerRunState(e, alive) {
        if (typeof e === 'undefined') {
            Logger.log("Trigger ID unknown.");
            return;
        }

        const key = CacheFinance.makeCacheKey("ALIVE", e.triggerUid);
        const shortCache = CacheService.getScriptCache();
        if (alive) {
            shortCache.put(key, "ALIVE");
        }
        else {
            if (shortCache.get(key) !== null) {
                shortCache.remove(key);
            }
        }
    }

    /**
     * Mark the job triggerID if not a valid trigger (in our job table.)
     * @returns {void}
     */
    validateTriggerIDs() {
        Logger.log("Starting validateTriggerIDs()");

        const validTriggerList = ScriptApp.getProjectTriggers();

        for (const job of this.jobs) {
            if (job.triggerID === "")
                continue;

            let good = false;
            for (const validID of validTriggerList) {
                if (job.triggerID === validID.getUniqueId()) {
                    good = true;
                    break;
                }
            }

            if (!good) {
                //  It could be running.
                if (CacheJobSettings.isDisabledTriggerStillRunning(job.triggerID)) {
                    continue;
                }

                Logger.log(`Invalid Trigger ID=${job.triggerID}`);
                job.triggerID = "";
            }
            else {
                Logger.log(`Valid Trigger ID=${job.triggerID}`);
            }
        }
    }

    /**
     * 
     * @param {String} triggerID 
     * @returns {Boolean}
     */
    static isDisabledTriggerStillRunning(triggerID) {
        const shortCache = CacheService.getScriptCache();

        const key = CacheFinance.makeCacheKey("ALIVE", triggerID);
        if (shortCache.get(key) !== null) {
            Logger.log(`Trigger ID=${triggerID} is RUNNING!`);
            return true;
        }

        return false;
    }

    /**
     * 
     * @param {any} triggerID 
     * @returns {void}
     */
    static deleteOldTrigger(triggerID) {
        const triggers = ScriptApp.getProjectTriggers();

        let triggerObject = null;
        for (const item of triggers) {
            if (item.getUniqueId() === triggerID)
                triggerObject = item;
        }

        if (triggerObject !== null) {
            Logger.log(`DELETING Trigger: ${triggerID}`);
            ScriptApp.deleteTrigger(triggerObject);
        }
        else {
            Logger.log(`Failed to locate trigger to delete: ${triggerID}`);
        }
    }

    /**
     * 
     * @param {Boolean} runAsap 
     * @returns {void}
     */
    createMissingTriggers(runAsap) {
        for (const job of this.jobs) {
            if (job.triggerID === "" && job.isValidJob()) {
                this.createTrigger(job, job.getMinutesToNextRun(runAsap) * 60);
            }
        }
    }

    /**
     * 
     * @param {CacheJob} job 
     * @param {Number} startAfterSeconds 
     * @returns {void}
     */
    createTrigger(job, startAfterSeconds) {
        if (job.timeout) {
            startAfterSeconds = 15;
        }

        const newTriggerID = ScriptApp
            .newTrigger('CacheFinanceTrigger')
            .timeBased()
            .after(startAfterSeconds * 1000)
            .create();
        job.triggerID = newTriggerID.getUniqueId();

        if (job.timeout) {
            const savedPartialResults = JSON.stringify(job.jobData);
            const key = CacheFinance.makeCacheKey("RESUMEJOB", job.triggerID);
            const shortCache = CacheService.getScriptCache();
            shortCache.put(key, savedPartialResults, 300);
            Logger.log(`Saving PARTIAL RESULTS.  Key=${key}. Items=${job.jobData.length}`);
        }

        this.updateFinanceCacheLegend();
        Logger.log(`New Trigger Created. ID=${job.triggerID}. Attrib=${job.attribute}. Start in ${startAfterSeconds} seconds.`);
    }

    /**
     * Update job status data to sheet.
     * @returns {void}
     */
    updateFinanceCacheLegend() {
        const legend = [];

        for (const job of this.jobs) {
            legend.push(job.toArray());
        }

        Logger.log(`Updating LEGEND.${legend}`);

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
                Logger.log(`Found My Job ID=${job.triggerID}`);
                myJob = job;
                break;
            }
        }

        if (myJob === null) {
            Logger.log(`This JOB ${runningScriptID} not found in legend.`);
        }
        else {
            //  Is this a continuation job?
            const key = CacheFinance.makeCacheKey("RESUMEJOB", myJob.triggerID);
            const shortCache = CacheService.getScriptCache();
            const resumeData = shortCache.get(key);
            if (resumeData !== null) {
                myJob.wasRestarted = true;
                myJob.jobData = JSON.parse(resumeData);
                Logger.log(`Resuming Job: ${key}. Items added=${myJob.jobData.length}`);
            }
            else {
                Logger.log(`New Job.  Key=${key}`);
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

        Logger.log(`Job Settings:  symbols=${this.symbolRange}. attribute=${this.attribute}. Out=${this.outputRange}. In=${this.defaultRange}. Minutes=${this.refreshMinutes}. Hours=${this.hours}. Days=${this.days}`);

        if (typeof this.hours === 'string')
            this.hours = this.hours.trim().toUpperCase();

        if (typeof this.days === 'string')
            this.days = this.days.trim().toUpperCase();

        this.extractRunDaysOfWeek();
        this.extractHoursOfDay();
    }

    /**
     * 
     * @returns {Boolean}
     */
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

    /**
     * 
     * @param {Boolean} value 
     */
    timedOut(value) {
        this.timeout = value;
    }

    /**
     * 
     * @returns {Boolean}
     */
    isRestarting() {
        return this.wasRestarted;
    }

    /**
     * Extract days that can be run from job legend into our job info.
     * @returns {void}
     */
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

        const dayNameIndex = this.DAYNAMES.indexOf(this.days);
        if (dayNameIndex !== -1) {
            this.dayNumbers[dayNameIndex] = true;
            return;
        }

        const singleItem = parseInt(this.days, 10);
        if (!isNaN(singleItem)) {
            this.dayNumbers[singleItem] = true;
            return;
        }

        Logger.log(`Job does not contain any days of the week to run. ${this.days}`);
    }

    /**
     * 
     * @returns {Boolean}
     */
    extractListDaysOfWeek() {
        if (this.days.indexOf(",") === -1) {
            return false;
        }

        let listValues = this.days.split(",");
        listValues = listValues.map(p => p.trim());

        for (const day of listValues) {
            let dayNameIndex = this.DAYNAMES.indexOf(day);
            if (dayNameIndex === -1) {
                dayNameIndex = parseInt(day, 10);
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

    /**
     * 
     * @returns {Boolean}
     */
    extractRangeDaysOfWeek() {
        if (this.days.indexOf("-") === -1) {
            return false;
        }

        let rangeValues = this.days.split("-");
        rangeValues = rangeValues.map(p => p.trim());

        let startDayNameIndex = this.DAYNAMES.indexOf(rangeValues[0]);
        let endDayNameIndex = this.DAYNAMES.indexOf(rangeValues[1]);
        if (startDayNameIndex === -1 || endDayNameIndex === -1) {
            startDayNameIndex = parseInt(rangeValues[0], 10);
            endDayNameIndex = parseInt(rangeValues[1], 10);

            if (isNaN(startDayNameIndex) || isNaN(endDayNameIndex)) {
                return false;
            }
        }

        let count = 0;
        for (let i = startDayNameIndex; count < 7; count++) {
            this.dayNumbers[i] = true;

            if (i === endDayNameIndex)
                break;

            i++;
            if (i > 6)
                i = 0;
        }

        return true;
    }

    /**
     * Parse HOURS set in legend into 24 hour true/false array.
     * @returns {void}
     */
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

        const singleItem = parseInt(this.hours, 10);
        if (!isNaN(singleItem)) {
            this.hourNumbers[singleItem] = true;
            return;
        }

        Logger.log(`This job does not contain any valid hours to run. ${this.hours}`);
    }

    /**
     * 
     * @returns {Boolean}
     */
    extractListHoursOfDay() {
        if (this.hours.indexOf(",") === -1) {
            return false;
        }

        let listValues = this.hours.split(",");
        listValues = listValues.map(p => p.trim());

        for (const hr of listValues) {
            let hourIndex = parseInt(hr, 10);
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

    /**
     * 
     * @returns {Boolean}
     */
    extractRangeHoursOfDay() {
        if (this.hours.indexOf("-") === -1) {
            return false;
        }

        let rangeValues = this.hours.split("-");
        rangeValues = rangeValues.map(p => p.trim());

        const startHourIndex = parseInt(rangeValues[0], 10);
        const endHourIndex = parseInt(rangeValues[1], 10);

        if (isNaN(startHourIndex) || isNaN(endHourIndex)) {
            return false;
        }

        if (startHourIndex < 0 || startHourIndex > 23 ||
            endHourIndex < 0 || endHourIndex > 23) {
            return false;
        }

        let count = 0;
        for (let i = startHourIndex; count < 24; count++) {
            this.hourNumbers[i] = true;

            if (i === endHourIndex)
                break;

            i++;
            if (i > 23)
                i = 0;
        }

        return true;
    }

    /**
     * 
     * @param {Boolean} runAsap 
     * @returns {Number}
     */
    getMinutesToNextRun(runAsap = false) {
        let minutes = this.refreshMinutes < 1 || runAsap ? 1 : this.refreshMinutes;

        // Get current date
        const startDateTime = new Date();
        let daysSinceStart = 0;
        const date = new Date();
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
        const dow = date.getDay();
        const status = this.dayNumbers[dow];

        return status;
    }

    /**
     * 
     * @param {Date} date 
     * @returns {Boolean}
     */
    canRunJobForTime(date) {
        const hour = date.getHours();
        const status = this.hourNumbers[hour];

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
        const hr = date.getHours();
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
        const row = [];
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
            Logger.log("* * *   Not time to run JOB   * * *");
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
     * @returns {void}
     */
    static getJobData(jobSettings) {
        const MAX_RUN_SECONDS = 300;
        let data = [];
        const symbolsNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.symbolRange);
        let googleValuesNamedRange = null;
        if (jobSettings.defaultRange !== "") {
            googleValuesNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.defaultRange);
        }

        Logger.log(`CacheTrigger: Symbols:${jobSettings.symbolRange}. Attribute:${jobSettings.attribute}. GoogleRange: ${jobSettings.defaultRange}`);

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

        if (googleValuesNamedRange !== null && symbols.length !== defaultData.length) {
            Logger.log(`Symbol Ranges and Google Values Ranges must be the same: ${jobSettings.symbolRange}. Len=${symbols.length}.  GoogleValues: ${jobSettings.defaultRange}. Len=${defaultData.length}`);
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
            const elapsed = (new Date().getTime() - start) / 1000;
            if (elapsed > MAX_RUN_SECONDS) {
                Logger.log("Max. Job Time reached.");
                jobSettings.timedOut(true);
                break;
            }

            let value = null;
            if (googleValuesNamedRange === null) {
                value = CACHEFINANCE(symbols[i][0], attribute);
            }
            else {
                value = CACHEFINANCE(symbols[i][0], attribute, defaultData[i][0]);
            }

            data.push([value]);
        }

        jobSettings.save(data);
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {Boolean}
     */
    static writeResults(jobSettings) {
        if (jobSettings.jobData === null || jobSettings.timeout)
            return false;

        Logger.log(`writeCacheResults:  START.  Data Len=${jobSettings.jobData.length}`);

        const outputNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.outputRange);

        if (outputNamedRange === null) {
            return false;
        }

        try {
            outputNamedRange.setValues(jobSettings.jobData);
        }
        catch (ex) {
            Logger.log(`Updating output range FAILED.  ${ex.toString()}`);
            return false;
        }

        Logger.log(`writeCacheResults:  END: ${jobSettings.outputRange}`);
        return true;
    }
}


/** @classdesc 
 * Stores settings for the SCRIPT.  Long term cache storage for small tables.  */
class ScriptSettings {      //  skipcq: JS-0128
    /**
     * For storing cache data for very long periods of time.
     */
    constructor() {
        this.scriptProperties = PropertiesService.getScriptProperties();
    }

    /**
     * Get script property using key.  If not found, returns null.
     * @param {String} propertyKey 
     * @returns {any}
     */
    get(propertyKey) {
        const myData = this.scriptProperties.getProperty(propertyKey);

        if (myData === null)
            return null;

        /** @type {PropertyData} */
        const myPropertyData = JSON.parse(myData);

        return PropertyData.getData(myPropertyData);
    }

    /**
     * Put data into our PROPERTY cache, which can be held for long periods of time.
     * @param {String} propertyKey - key to finding property data.
     * @param {any} propertyData - value.  Any object can be saved..
     * @param {Number} daysToHold - number of days to hold before item is expired.
     */
    put(propertyKey, propertyData, daysToHold = 1) {
        //  Create our object with an expiry time.
        const objData = new PropertyData(propertyData, daysToHold);

        //  Our property needs to be a string
        const jsonData = JSON.stringify(objData);

        try {
            this.scriptProperties.setProperty(propertyKey, jsonData);
        }
        catch (ex) {
            throw new Error("Cache Limit Exceeded.  Long cache times have limited storage available.  Only cache small tables for long periods.");
        }
    }

    /**
     * 
     * @param {Object} propertyDataObject 
     * @param {Number} daysToHold 
     */
    putAll(propertyDataObject, daysToHold = 1) {
        const keys = Object.keys(propertyDataObject);
        keys.forEach(key => this.put(key, propertyDataObject[key], daysToHold));
    }

    /**
     * Removes script settings that have expired.
     * @param {Boolean} deleteAll - true - removes ALL script settings regardless of expiry time.
     */
    expire(deleteAll) {
        const allKeys = this.scriptProperties.getKeys();

        for (const key of allKeys) {
            const myData = this.scriptProperties.getProperty(key);

            if (myData !== null) {
                let propertyValue = null;
                try {
                    propertyValue = JSON.parse(myData);
                }
                catch (e) {
                    Logger.log(`Script property data is not JSON. key=${key}`);
                }

                if (propertyValue !== null && (PropertyData.isExpired(propertyValue) || deleteAll)) {
                    this.scriptProperties.deleteProperty(key);
                    Logger.log(`Removing expired SCRIPT PROPERTY: key=${key}`);
                }
            }
        }
    }
}

/** Converts data into JSON for getting/setting in ScriptSettings. */
class PropertyData {
    /**
     * 
     * @param {any} propertyData 
     * @param {Number} daysToHold 
     */
    constructor(propertyData, daysToHold) {
        const someDate = new Date();

        /** @property {String} */
        this.myData = JSON.stringify(propertyData);
        /** @property {Date} */
        this.expiry = someDate.setMinutes(someDate.getMinutes() + daysToHold * 1440);
    }

    /**
     * 
     * @param {PropertyData} obj 
     * @returns {any}
     */
    static getData(obj) {
        let value = null;
        try {
            if (!PropertyData.isExpired(obj)) {
                value = JSON.parse(obj.myData);
            }
        }
        catch (ex) {
            Logger.log(`Invalid property value.  Not JSON: ${ex.toString()}`);
        }

        return value;
    }

    /**
     * 
     * @param {PropertyData} obj 
     * @returns {Boolean}
     */
    static isExpired(obj) {
        const someDate = new Date();
        const expiryDate = new Date(obj.expiry);
        return (expiryDate.getTime() < someDate.getTime())
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
        let data = new StockAttributes();

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
                Logger.log(`3'rd Party FINANCE attribute not supported: ${attribute}`);
                break;
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

    /**
     * 
     * @param {String} symbol 
     * @returns {StockAttributes}
     */
    static getName(symbol) {
        /** @type {StockAttributes} */
        let data = GlobeAndMail.getInfo(symbol);

        if (data.stockName === null)
            data = TdMarketResearch.getInfo(symbol, "ETF");

        if (data.stockName === null)
            data = TdMarketResearch.getInfo(symbol, "STOCK");

        return data;
    }

    /**
     * 
     * @param {String} symbol 
     * @returns {String}
     */
    static getTickerCountryCode(symbol) {
        const colon = symbol.indexOf(":");
        let exchange = "";
        let countryCode = "";

        if (colon < 0) {
            return countryCode;
        }

        const parts = symbol.split(":");
        exchange = parts[0].toUpperCase();

        switch(exchange) {
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
                countryCode = "us";
                break; 
        }

        return countryCode;
    }
}

/**
 * @classdesc Lookup for TD website.
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
            URL = `https://marketsandresearch.td.com/tdwca/Public/ETFsProfile/Summary/${ThirdPartyFinance.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;
        else
            URL = `https://marketsandresearch.td.com/tdwca/Public/Stocks/Overview/${ThirdPartyFinance.getTickerCountryCode(symbol)}/${TdMarketResearch.getTicker(symbol)}`;

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
            symbol = `${symbol.substr(0, dash)}.PR.${symbol.substr(dash + 1)}`;
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

/**
 * Returns a diagnostic list of 3rd party stock lookup info.
 * @returns {any[][]}
 */
function cacheFinanceTest() {
    const tester = new CacheFinanceTest();

    return tester.execute();
}

/**
 * @classdesc executes 3rd party data lookup tests.
 */
class CacheFinanceTest {
    constructor() {
        this.cacheTestRun = new CacheFinanceTestRun();
    }

    /**
     * Run the tests.
     * @returns {any[][]}
     */
    execute() {
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "NYSEARCA:SHYG");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "badSymbol");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "ZTL");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "TSE:MEG");
        this.cacheTestRun.run("Yahoo", YahooFinance.getInfo, "NASDAQ:MSFT");
        
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "NYSEARCA:SHYG");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "ZTL");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:ZTL");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:DFN-A");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "ZTL", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "TSE:DFN-A", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "NASDAQ:MSFT", "STOCK");
        this.cacheTestRun.run("TD", TdMarketResearch.getInfo, "badSymbol");

        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NYSEARCA:SHYG");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:FTN-A");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:HBF.B");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:MEG");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "TSE:ZTL");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "NASDAQ:MSFT");
        this.cacheTestRun.run("GlobeAndMail", GlobeAndMail.getInfo, "badSymbol");

        return this.cacheTestRun.getTestRunResults();
    }
}

/**
 * @classdesc track test results for 3rd party site tests.
 */
class CacheFinanceTestRun {
    constructor() {
        /** @type {CacheFinanceTestStatus[]} */
        this.testRuns = [];
    }

    /**
     * Run one test.
     * @param {String} serviceName 
     * @param {*} func 
     * @param {String} symbol 
     */
    run(serviceName, func, symbol, type = "ETF") {
        const result = new CacheFinanceTestStatus(serviceName, symbol);
        try {
            /** @type {StockAttributes} */
            const data = func(symbol, type);
            result.setStatus("ok");
            result.setStockAttributes(data);
            result.setTypeLookup(type);

            if (data.stockName === null && data.stockPrice === null && data.yieldPct === null) {
                result.setStatus("Not Found!")
            }
        }
        catch(ex) {
            result.setStatus("Error");
        }
        result.finishTimer();

        this.testRuns.push(result);
    } 
    
    /**
     * Return results for all tests.
     * @returns {any[][]}
     */
    getTestRunResults() {
        const resultTable = [];

        /** @type {any[]} */
        let row = ["Service", "Symbol", "Status", "Price", "Yield", "Name", "Type", "Run Time(ms)"];
        resultTable.push(row);

        for (const testRun of this.testRuns) {
            row = [];

            row.push(testRun.serviceName);
            row.push(testRun.symbol);
            row.push(testRun.status);
            row.push(testRun.stockAttributes.stockPrice);
            row.push(testRun.stockAttributes.yieldPct);
            row.push(testRun.stockAttributes.stockName);
            row.push(testRun.typeLookup);
            row.push(testRun.runTime);

            resultTable.push(row);
        }

        return resultTable;
    }
}

/**
 * Individual test results and tracking.
 */
class CacheFinanceTestStatus {
    constructor (serviceName="", symbol="") {
        this._serviceName = serviceName;
        this._symbol = symbol;
        this._stockAttributes = new StockAttributes();
        this._startTime = Date.now()
        this._typeLookup = "";
        this._runTime = 0;
    }

    get serviceName() {
        return this._serviceName;
    }
    get symbol () {
        return this._symbol;
    }
    get value() {
        return this._value;
    }
    get status() {
        return this._status;
    }
    get runTime() {
        return this._runTime;
    }
    get stockAttributes() {
        return this._stockAttributes;
    }
    get typeLookup() {
        return this._typeLookup;
    }

    /**
     * 
     * @param {String} val 
     * @returns {CacheFinanceTestStatus}
     */
    setServiceName(val) {
        this._serviceName = val;
        return this;
    }

    /**
     * 
     * @param {String} val 
     * @returns {CacheFinanceTestStatus}
     */
    setSymbol(val) {
        this._symbol = val;
        return this;
    }

    /**
     * 
     * @param {any} val 
     * @returns {CacheFinanceTestStatus}
     */
    setValue(val) {
        this._value = val;
        return this;
    }

    /**
     * 
     * @param {String} val 
     * @returns {CacheFinanceTestStatus}
     */
    setStatus(val) {
        this._status = val;
        return this;
    }

    /**
     * 
     * @param {String} val 
     * @returns {CacheFinanceTestStatus}
     */
    setTypeLookup(val) {
        this._typeLookup = val;
        return this;
    }

    /**
     * 
     * @param {StockAttributes} val 
     * @returns {CacheFinanceTestStatus}
     */
    setStockAttributes(val) {
        this._stockAttributes = val;
        return this;
    }

    /**
     * 
     * @returns {CacheFinanceTestStatus}
     */
    finishTimer() {
        this._runTime = Date.now() - this._startTime;
        return this;
    }
}



