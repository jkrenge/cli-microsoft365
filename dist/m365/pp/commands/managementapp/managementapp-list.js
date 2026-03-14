import { odata } from '../../../../utils/odata.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpManagementAppListCommand extends PowerPlatformCommand {
    get name() {
        return commands.MANAGEMENTAPP_LIST;
    }
    get description() {
        return 'Lists management applications for Power Platform';
    }
    async commandAction(logger) {
        const endpoint = `${this.resource}/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01`;
        try {
            const managementApps = await odata.getAllItems(endpoint);
            await logger.log(managementApps);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PpManagementAppListCommand();
//# sourceMappingURL=managementapp-list.js.map