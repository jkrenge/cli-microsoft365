import request from '../../../../request.js';
import SpoCommand from '../../../base/SpoCommand.js';
export class SpoAppBaseCommand extends SpoCommand {
    async getAppCatalogSiteUrl(logger, authSiteUrl, args) {
        if (args.options.appCatalogScope === 'sitecollection') {
            // trim trailing slashes if there are any
            const appCatalogUrl = args.options.appCatalogUrl.replace(/\/$/, '');
            const appCatalogUrlChunks = appCatalogUrl.split('/');
            // Trim the last part of the URL if it ends on '/appcatalog', but don't trim it if the site URL is called like that (/sites/appcatalog).
            if (appCatalogUrl.toLowerCase().endsWith('/appcatalog') && appCatalogUrlChunks.length !== 5) {
                return appCatalogUrl.substring(0, appCatalogUrl.lastIndexOf('/'));
            }
        }
        if (args.options.appCatalogUrl) {
            return args.options.appCatalogUrl.replace(/\/$/, '');
        }
        if (this.verbose) {
            await logger.logToStderr('Getting tenant app catalog url...');
        }
        const requestOptions = {
            url: `${authSiteUrl}/_api/SP_TenantSettings_Current`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (response.CorporateCatalogUrl) {
            return response.CorporateCatalogUrl;
        }
        throw new Error('Tenant app catalog is not configured.');
    }
}
//# sourceMappingURL=SpoAppBaseCommand.js.map