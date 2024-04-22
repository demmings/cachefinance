/*  *** DEBUG START ***
//  Remove comments for testing in NODE

import { CACHEFINANCE, CacheFinance, GOOGLEFINANCE_PARAM_NOT_USED } from "./CacheFinance";

class Logger {
    static log(msg) {
        console.log(msg);
    }
}
//  *** DEBUG END ***/

/**
 * Run test manually within Sheets.
 */
function testCacheTriggerJob() {                // skipcq: JS-0128
    const jobParms = ["Stocks!A8:A140",
        "PRICE",
        "Stocks!C8:C140",
        "Stocks!CB8:CB140",
        1,
        "*",
        "MON-SUN",
        ""];

    const jobSettings = new CacheJob(jobParms);
    CacheTrigger.getJobData(jobSettings);
}

/**
 * Run test manually within Sheets.
 */
function testBulkCache() {                      // skipcq: JS-0128
    const symbols = [["ABC"], ["DEF"], ["GHI"], ["JKL"], ["MNO"]];
    const data = [11.1, 22.2, 33.3, 44.4, 55.5];

    CacheTrigger.bulkCachePut(symbols, "PRICE", data);
    const cacheData = CacheTrigger.bulkCacheGet(symbols, "PRICE");

    if (JSON.stringify(data) !== JSON.stringify(cacheData)) {
        Logger.log("BULK Cache TEST Fail.");
    }

    Logger.log("BULK CACHE TEST Success");
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
 * @classdesc Manage TRIGGERS that run and update stock/etf data from the web.
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
        else if (shortCache.get(key) !== null) {
            shortCache.remove(key);
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
        this.maxRunSeconds = 0;
        this.startJobTime = null;

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
     * @param {Number} maxRunSeconds
     * @returns {any[]}
     */
    startJobRunTimer(maxRunSeconds) {
        let data = [];
        if (this.isRestarting()) {
            data = this.jobData;
        }

        this.timedOut(false);
        this.maxRunSeconds = maxRunSeconds;
        this.startJobTime = new Date().getTime();

        return data;
    }

    /**
     * 
     * @returns {Boolean}
     */
    hasExceededJobRunTime() {
        const elapsed = (new Date().getTime() - this.startJobTime) / 1000;
        if (elapsed > this.maxRunSeconds) {
            Logger.log("Max. Job Time reached.");
            this.timedOut(true);
            return true;
        }

        return false;
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
            this.hourNumbers[i] = this.hours === '*';

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
     * @param {CacheJob} jobSettings - results also returned in object.
     * @returns {void}
     */
    static getJobData(jobSettings) {
        const MAX_RUN_SECONDS = 300;
        const attribute = jobSettings.attribute;
        const symbols = CacheTrigger.getSymbols(jobSettings);
        const defaultData = CacheTrigger.getDefaultData(jobSettings);
        const bulkDataCache = CacheTrigger.bulkCacheGet(symbols, attribute);

        Logger.log(`CacheTrigger: Symbols:${jobSettings.symbolRange}. Attribute:${jobSettings.attribute}. GoogleRange: ${jobSettings.defaultRange}`);

        if (CacheTrigger.isDefaultDataProvided(jobSettings) && symbols.length !== defaultData.length) {
            throw new Error(`Symbol Ranges and Google Values Ranges must be the same: ${jobSettings.symbolRange}. Len=${symbols.length}.  GoogleValues: ${jobSettings.defaultRange}. Len=${defaultData.length}`);
        }

        //  Start job timer and load previously loaded (and not finished) job data.
        const data = jobSettings.startJobRunTimer(MAX_RUN_SECONDS);
        const startingSymbol = data.length;

        for (let i = startingSymbol; i < symbols.length; i++) {
            if (jobSettings.hasExceededJobRunTime()) {
                break;
            }

            let value = null;
            if (CacheTrigger.isDefaultDataProvided(jobSettings)) {
                value = CacheFinance.getFinanceData(symbols[i][0], attribute, defaultData[i][0], bulkDataCache[i]);
            }
            else {
                value = CacheFinance.getFinanceData(symbols[i][0], attribute, GOOGLEFINANCE_PARAM_NOT_USED);
            }

            data.push([value]);
        }

        jobSettings.save(data);
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {any[][]}
     */
    static getSymbols(jobSettings) {
        const symbolsNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.symbolRange);
        if (symbolsNamedRange === null) {
            throw new Error(`Failed to read data from SYMBOLS range =${jobSettings.symbolRange}`);
        }

        return symbolsNamedRange.getValues();
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {any[][]}
     */
    static getDefaultData(jobSettings) {
        let defaultData = [];

        if (jobSettings.defaultRange !== "") {
            const symbolsNamedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(jobSettings.defaultRange);
            if (symbolsNamedRange === null) {
                throw new Error(`Failed to read data from DEFAULT DATA range =${jobSettings.defaultRange}`);
            }

            defaultData = symbolsNamedRange.getValues();
        }

        return defaultData;
    }

    /**
     * 
     * @param {CacheJob} jobSettings 
     * @returns {Boolean}
     */
    static isDefaultDataProvided(jobSettings) {
        return jobSettings.defaultRange !== "";
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

    /**
     * 
     * @param {any[][]} symbols 
     * @param {String} attribute 
     * @returns {any[]} 
     */
    static bulkCacheGet(symbols, attribute) {
        const cacheKeyList = CacheTrigger.createCacheKeyList(symbols, attribute);
        return CacheTrigger.getFinanceValuesFromShortCache(cacheKeyList);
    }

    /**
     * 
     * @param {any[][]} symbols 
     * @param {String} attribute 
     * @param {any[]} newCacheData 
     */
    static bulkCachePut(symbols, attribute, newCacheData) {
        const cacheKeyList = CacheTrigger.createCacheKeyList(symbols, attribute);
        CacheTrigger.putFinanceValuesIntoShortCache(cacheKeyList, newCacheData);
    }

    /**
     * 
     * @param {any[][]} symbols 
     * @param {String} attribute 
     * @returns {any[]}
     */
    static createCacheKeyList(symbols, attribute) {
        const cacheKeyList = [];
        for (const symbol of symbols) {
            attribute = attribute.toUpperCase().trim();
            const upperCaseSymbol = symbol[0].toUpperCase();
            const cacheKey = CacheFinance.makeCacheKey(upperCaseSymbol, attribute.toUpperCase().trim());
            cacheKeyList.push(cacheKey);
        }

        return cacheKeyList;
    }

    /**
     * 
     * @param {String[]} cacheKeys 
     * @returns {any[]}
     */
    static getFinanceValuesFromShortCache(cacheKeys) {
        const shortCache = CacheService.getScriptCache();

        //  Object with key/value pairs for all items found in cache.
        const data = shortCache.getAll(cacheKeys);
        const cachedDataList = [];

        for (const key of cacheKeys) {
            let parsedData = null;
            if (typeof data[key] !== 'undefined') {
                parsedData = JSON.parse(data[key]);
            }
            cachedDataList.push(parsedData);
        }

        return cachedDataList;
    }

    /**
     * Puts list of data into cache using one API call.  Data is converted to JSON before it is updated.
     * @param {String[]} cacheKeys 
     * @param {any[]} newCacheData 
     */
    static putFinanceValuesIntoShortCache(cacheKeys, newCacheData) {
        const bulkData = {};

        for (let i = 0; i < cacheKeys.length; i++) {
            bulkData[cacheKeys[i]] = JSON.stringify(newCacheData[i]);
        }

        const shortCache = CacheService.getScriptCache();
        shortCache.putAll(bulkData, 21600);
    }
}
