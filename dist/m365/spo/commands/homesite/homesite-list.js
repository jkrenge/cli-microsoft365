import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHomeSiteListCommand extends SpoCommand {
    get name() {
        return commands.HOMESITE_LIST;
    }
    get description() {
        return 'Lists all home sites';
    }
    defaultProperties() {
        return ['Url', 'Title'];
    }
    async commandAction(logger) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving all home sites...`);
            }
            const res = await odata.getAllItems(`${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoHomeSiteListCommand();
//# sourceMappingURL=homesite-list.js.map