import { cli } from "../cli/cli.js";
import request from "../request.js";
import { formatting } from "./formatting.js";
import { odata } from "./odata.js";
const powerPlatformResource = 'https://api.bap.microsoft.com';
export const powerPlatform = {
    async getDynamicsInstanceApiUrl(environment, asAdmin) {
        let url;
        if (asAdmin) {
            url = `${powerPlatformResource}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/${formatting.encodeQueryParameter(environment)}`;
        }
        else {
            url = `${powerPlatformResource}/providers/Microsoft.BusinessAppPlatform/environments/${formatting.encodeQueryParameter(environment)}`;
        }
        const requestOptions = {
            url: `${url}?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const response = await request.get(requestOptions);
            return response.properties.linkedEnvironmentMetadata.instanceApiUrl;
        }
        catch (ex) {
            throw Error(`The environment '${environment}' could not be retrieved. See the inner exception for more details: ${ex.message}`);
        }
    },
    async getWebsiteById(environment, id) {
        const requestOptions = {
            url: `https://api.powerplatform.com/powerpages/environments/${environment}/websites/${id}?api-version=2022-03-01-preview`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const response = await request.get(requestOptions);
            return response;
        }
        catch {
            throw Error(`The specified Power Page website with id '${id}' does not exist.`);
        }
    },
    async getWebsiteByName(environment, websiteName) {
        const response = await odata.getAllItems(`https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`);
        const items = response.filter(response => response.name === websiteName);
        if (items.length === 0) {
            throw Error(`The specified Power Page website '${websiteName}' does not exist.`);
        }
        if (items.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('websiteUrl', items);
            return cli.handleMultipleResultsFound(`Multiple Power Page websites with name '${websiteName}' found.`, resultAsKeyValuePair);
        }
        return items[0];
    },
    async getWebsiteByUrl(environment, url) {
        const response = await odata.getAllItems(`https://api.powerplatform.com/powerpages/environments/${environment}/websites?api-version=2022-03-01-preview`);
        const items = response.filter(response => response.websiteUrl === url);
        if (items.length === 0) {
            throw Error(`The specified Power Page website with url '${url}' does not exist.`);
        }
        return items[0];
    },
    /**
     * Get a card by name
     * Returns a card object
     * @param dynamicsApiUrl The dynamics api url of the environment
     * @param name The name of the card
     * @param logger The logger object
     * @param verbose Set for verbose logging
     */
    async getCardByName(dynamicsApiUrl, name) {
        const requestOptions = {
            url: `${dynamicsApiUrl}/api/data/v9.1/cards?$filter=name eq '${name}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const result = await request.get(requestOptions);
        if (result.value.length === 0) {
            throw Error(`The specified card '${name}' does not exist.`);
        }
        if (result.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('cardid', result.value);
            return cli.handleMultipleResultsFound(`Multiple cards with name '${name}' found.`, resultAsKeyValuePair);
        }
        return result.value[0];
    },
    /**
   * Get a solution by name
   * Returns the solution object
   * @param dynamicsApiUrl The dynamics api url of the environment
   * @param name The name of the solution
   */
    async getSolutionByName(dynamicsApiUrl, name) {
        const requestOptions = {
            url: `${dynamicsApiUrl}/api/data/v9.0/solutions?$filter=isvisible eq true and uniquename eq '${name}'&$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const result = await request.get(requestOptions);
        if (result.value.length === 0) {
            throw Error(`The specified solution '${name}' does not exist.`);
        }
        if (result.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('solutionid', result.value);
            return cli.handleMultipleResultsFound(`Multiple solutions with name '${name}' found.`, resultAsKeyValuePair);
        }
        return result.value[0];
    }
};
//# sourceMappingURL=powerPlatform.js.map