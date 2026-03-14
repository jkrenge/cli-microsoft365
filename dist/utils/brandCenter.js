import request from '../request.js';
import { spo } from './spo.js';
export const brandCenter = {
    /**
     * Gets the brand center configuration for the specified site
     * @param logger Logger instance for verbose output
     * @param debug Debug flag for detailed logging
     * @returns Promise<BrandCenterConfiguration> Brand center configuration object
     */
    async getBrandCenterConfiguration(logger, debug = false) {
        if (debug) {
            await logger.logToStderr(`Retrieving brand center configuration...`);
        }
        const spoAdminUrl = await spo.getSpoAdminUrl(logger, debug);
        const brandConfigRequestOptions = {
            url: `${spoAdminUrl}/_api/SPO.Tenant/GetBrandCenterConfiguration`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const brandConfig = await request.get(brandConfigRequestOptions);
        if (debug) {
            await logger.logToStderr(`Successfully retrieved brand center configuration`);
        }
        return brandConfig;
    }
};
//# sourceMappingURL=brandCenter.js.map