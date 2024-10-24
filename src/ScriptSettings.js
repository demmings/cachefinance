/*  *** DEBUG START ***
//  Remove comments for testing in NODE
export { ScriptSettings, PropertyData };
import { CacheFinance } from "../CacheFinance.js";
import { PropertiesService } from "../GasMocks.js";

class Logger {
    static log(msg) {
        console.log(msg);
    }
}

//  *** DEBUG END ***/

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

        if (PropertyData.isExpired(myPropertyData))
        {
            this.delete(propertyKey);
            return null;
        }

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
     * @param {Object} propertyDataObject 
     * @param {Number} daysToHold 
     */
    putAll(propertyDataObject, daysToHold = 1) {
        const keys = Object.keys(propertyDataObject);
        keys.forEach(key => this.put(key, propertyDataObject[key], daysToHold));
    }

    /**
     * Puts list of data into cache using one API call.  Data is converted to JSON before it is updated.
     * @param {String[]} cacheKeys 
     * @param {any[]} newCacheData 
     * @param {Number} daysToHold
     */
    static putAllKeysWithData(cacheKeys, newCacheData, daysToHold = 7) {
        const bulkData = {};

        for (let i = 0; i < cacheKeys.length; i++) {
            //  Create our object with an expiry time.
            const objData = new PropertyData(newCacheData[i], daysToHold);

            //  Our property needs to be a string
            bulkData[cacheKeys[i]] = JSON.stringify(objData);
        }

        PropertiesService.getScriptProperties().setProperties(bulkData);
    }

    /**
     * Returns ALL cached data for each key value requested. 
     * Only 1 API call is made, so much faster than retrieving single values.
     * @param {String[]} cacheKeys 
     * @returns {any[]}
     */
    static getAll(cacheKeys) {
        const values = [];

        if (cacheKeys.length === 0) {
            return values;
        }
        
        const allProperties = PropertiesService.getScriptProperties().getProperties();

        //  Removing properties is very slow, so remove only 1 at a time.  This is enough as this function is called frequently.
        ScriptSettings.expire(false, 1, allProperties);

        for (const key of cacheKeys) {
            const myData = allProperties[key];

            if (typeof myData === 'undefined') {
                values.push(null);
            }
            else {
                /** @type {PropertyData} */
                const myPropertyData = JSON.parse(myData);

                if (PropertyData.isExpired(myPropertyData)) {
                    values.push(null);
                    PropertiesService.getScriptProperties().deleteProperty(key);
                    Logger.log(`Delete expired Script Property Key=${key}`);
                }
                else {
                    values.push(PropertyData.getData(myPropertyData));
                }
            }
        }

        return values;
    }

    /**
     * Removes script settings that have expired.
     * @param {Boolean} deleteAll - true - removes ALL script settings regardless of expiry time.
     * @param {Number} maxDelete - maximum number of items to delete that are expired.
     * @param {Object} allPropertiesObject - All properties already loaded.  If null, will load iteself.
     */
    static expire(deleteAll, maxDelete = 999, allPropertiesObject = null) {
        const allProperties = allPropertiesObject === null ? PropertiesService.getScriptProperties().getProperties() : allPropertiesObject;
        const allKeys = Object.keys(allProperties);
        let deleteCount = 0;

        for (const key of allKeys) {
            let propertyValue = null;
            try {
                propertyValue = JSON.parse(allProperties[key]);
            }
            catch (e) {
                //  A property that is NOT cached by CACHEFINANCE
                continue;
            }

            const propertyOfThisApplication = propertyValue?.expiry !== undefined;

            if (propertyOfThisApplication && (PropertyData.isExpired(propertyValue) || deleteAll)) {
                PropertiesService.getScriptProperties().deleteProperty(key);
                delete allProperties[key];

                //  There is no way to iterate existing from 'short' cache, so we assume there is a
                //  matching short cache entry and attempt to delete.
                CacheFinance.deleteFromShortCache(key);

                Logger.log(`Removing expired SCRIPT PROPERTY: key=${key}`);

                deleteCount++;
            }

            if (deleteCount >= maxDelete) {
                return;
            }
        }
    }

    /**
     * Delete a specific key in script properties.
     * @param {String} key 
     */
    delete(key) {
        if (this.scriptProperties.getProperty(key) !== null) {
            this.scriptProperties.deleteProperty(key);
        }
    }
}

/**
 * @classdesc Converts data into JSON for getting/setting in ScriptSettings.
 */
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
     * @param {PropertyData} obj 
     * @returns {any}
     */
    static getData(obj) {
        let value = null;
        try {
            value = JSON.parse(obj.myData);
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