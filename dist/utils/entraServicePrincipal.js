import { odata } from './odata.js';
import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';
import request from '../request.js';
export const entraServicePrincipal = {
    /**
     * Get service principal by its appId
     * @param appId App id.
     * @param properties Comma-separated list of properties to include in the response.
     * @returns The service principal.
     * @throws Error when service principal was not found.
     */
    async getServicePrincipalByAppId(appId, properties) {
        let url = `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '${appId}'`;
        if (properties) {
            url += `&$select=${properties}`;
        }
        const apps = await odata.getAllItems(url);
        if (apps.length === 0) {
            throw new Error(`Service principal with appId '${appId}' not found in Microsoft Entra ID.`);
        }
        return apps[0];
    },
    /**
     * Get service principal by its name
     * @param appName Service principal name.
     * @param properties Comma-separated list of properties to include in the response.
     * @returns The service principal.
     * @throws Error when service principal was not found.
     */
    async getServicePrincipalByAppName(appName, properties) {
        let url = `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'`;
        if (properties) {
            url += `&$select=${properties}`;
        }
        const apps = await odata.getAllItems(url);
        if (apps.length === 0) {
            throw new Error(`Service principal with name '${appName}' not found in Microsoft Entra ID.`);
        }
        if (apps.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', apps);
            return await cli.handleMultipleResultsFound(`Multiple service principals with name '${appName}' found in Microsoft Entra ID.`, resultAsKeyValuePair);
        }
        return apps[0];
    },
    /**
     * Get all available service principals.
     * @param properties Comma-separated list of properties to include in the response.
     */
    async getServicePrincipals(properties) {
        let url = `https://graph.microsoft.com/v1.0/servicePrincipals`;
        if (properties) {
            url += `?$select=${properties}`;
        }
        return odata.getAllItems(url);
    },
    /**
     * Create a new service principal for the specified application.
     * @param appId Application ID of the application for which to create a service principal.
     * @returns The created service principal.
     */
    async createServicePrincipal(appId) {
        const url = `https://graph.microsoft.com/v1.0/servicePrincipals`;
        const requestOptions = {
            url: url,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
            },
            data: {
                appId
            },
            responseType: 'json'
        };
        return await request.post(requestOptions);
    }
};
//# sourceMappingURL=entraServicePrincipal.js.map