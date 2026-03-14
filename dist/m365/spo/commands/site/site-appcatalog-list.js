import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    excludeDeletedSites: z.boolean().optional()
});
class SpoSiteAppCatalogListCommand extends SpoCommand {
    get name() {
        return commands.SITE_APPCATALOG_LIST;
    }
    get description() {
        return 'List all site collection app catalogs within the tenant';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['AbsoluteUrl', 'SiteID'];
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr('Retrieving site collection app catalogs...');
            }
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            let appCatalogs = await odata.getAllItems(`${spoUrl}/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites`);
            if (args.options.excludeDeletedSites) {
                if (this.verbose) {
                    await logger.logToStderr('Excluding inaccessible sites from the results...');
                }
                const activeAppCatalogs = [];
                for (const appCatalog of appCatalogs) {
                    try {
                        await spo.getWeb(appCatalog.AbsoluteUrl, logger, this.verbose);
                        activeAppCatalogs.push(appCatalog);
                    }
                    catch (error) {
                        if (this.debug) {
                            await logger.logToStderr(error);
                        }
                        if (error.status === 404 || error.status === 403) {
                            if (this.verbose) {
                                await logger.logToStderr(`Site at '${appCatalog.AbsoluteUrl}' is inaccessible. Excluding from results...`);
                            }
                            continue;
                        }
                        throw error;
                    }
                }
                appCatalogs = activeAppCatalogs;
            }
            await logger.log(appCatalogs);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoSiteAppCatalogListCommand();
//# sourceMappingURL=site-appcatalog-list.js.map