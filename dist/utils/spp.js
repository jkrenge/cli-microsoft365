import request from '../request.js';
import { formatting } from './formatting.js';
export const spp = {
    /**
     * Asserts whether the specified site is a content center
     * @param siteUrl The URL of the site to check
     * @param logger Logger instance
     * @param verbose Whether to log verbose messages
     * @throws error when site is not a content center.
     */
    async assertSiteIsContentCenter(siteUrl, logger, verbose) {
        if (verbose) {
            await logger.log(`Checking if '${siteUrl}' is a valid content center site...`);
        }
        const requestOptions = {
            url: `${siteUrl}/_api/web?$select=WebTemplateConfiguration`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (response.WebTemplateConfiguration !== 'CONTENTCTR#0') {
            throw Error(`${siteUrl} is not a content site.`);
        }
    },
    /**
     * Gets a SharePoint Premium model by title
     * @param contentCenterUrl a content center site URL
     * @param title model title
     * @param logger Logger instance
     * @param verbose Whether to log verbose messages
     * @returns SharePoint Premium model
     */
    async getModelByTitle(contentCenterUrl, title, logger, verbose) {
        if (verbose) {
            await logger.log(`Retrieving model information...`);
        }
        let requestTitle = title.toLowerCase();
        if (!requestTitle.endsWith('.classifier')) {
            requestTitle += '.classifier';
        }
        const requestUrl = `${contentCenterUrl}/_api/machinelearning/models/getbytitle('${formatting.encodeQueryParameter(requestTitle)}')`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const result = await request.get(requestOptions);
        if (result['odata.null'] === true) {
            throw Error(`Model '${title}' was not found.`);
        }
        return result;
    },
    /**
     * Gets a SharePoint Premium model by unique id
     * @param contentCenterUrl a content center site URL
     * @param id model unique id
     * @param logger Logger instance
     * @param verbose Whether to log verbose messages
     * @returns SharePoint Premium model
     */
    async getModelById(contentCenterUrl, id, logger, verbose) {
        if (verbose) {
            await logger.log(`Retrieving model information...`);
        }
        const requestUrl = `${contentCenterUrl}/_api/machinelearning/models/getbyuniqueid('${id}')`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return await request.get(requestOptions);
    }
};
//# sourceMappingURL=spp.js.map