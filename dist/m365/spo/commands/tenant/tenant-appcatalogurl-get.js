import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantAppCatalogUrlGetCommand extends SpoCommand {
    get name() {
        return commands.TENANT_APPCATALOGURL_GET;
    }
    get description() {
        return 'Gets the URL of the tenant app catalog';
    }
    async commandAction(logger) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestOptions = {
                url: `${spoUrl}/_api/SP_TenantSettings_Current`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                }
            };
            const res = await request.get(requestOptions);
            const json = JSON.parse(res);
            if (json.CorporateCatalogUrl) {
                await logger.log(json.CorporateCatalogUrl);
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr("Tenant app catalog is not configured.");
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoTenantAppCatalogUrlGetCommand();
//# sourceMappingURL=tenant-appcatalogurl-get.js.map