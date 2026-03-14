import { CommandError } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantCommandSetListCommand extends SpoCommand {
    get name() {
        return commands.TENANT_COMMANDSET_LIST;
    }
    get description() {
        return 'Retrieves a list of ListView Command Sets that are installed tenant-wide';
    }
    defaultProperties() {
        return ['Title', 'TenantWideExtensionComponentId', 'TenantWideExtensionListTemplate'];
    }
    async commandAction(logger) {
        const appCatalogUrl = await spo.getTenantAppCatalogUrl(logger, this.debug);
        if (!appCatalogUrl) {
            throw new CommandError('No app catalog URL found');
        }
        if (this.verbose) {
            await logger.logToStderr('Retrieving a list of ListView Command Sets that are installed tenant-wide');
        }
        const listServerRelativeUrl = urlUtil.getServerRelativePath(appCatalogUrl, '/lists/TenantWideExtensions');
        try {
            const listItems = await odata.getAllItems(`${appCatalogUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet')`);
            listItems.forEach(i => delete i.ID);
            await logger.log(listItems);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoTenantCommandSetListCommand();
//# sourceMappingURL=tenant-commandset-list.js.map