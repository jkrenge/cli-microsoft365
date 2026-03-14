import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
export const options = globalOptionsZod.strict();
class PpTenantSettingsListCommand extends PowerPlatformCommand {
    get name() {
        return commands.TENANT_SETTINGS_LIST;
    }
    get description() {
        return 'Lists the global Power Platform tenant settings';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['disableCapacityAllocationByEnvironmentAdmins', 'disableEnvironmentCreationByNonAdminUsers', 'disableNPSCommentsReachout', 'disablePortalsCreationByNonAdminUsers', 'disableSupportTicketsVisibleByAllUsers', 'disableSurveyFeedback', 'disableTrialEnvironmentCreationByNonAdminUsers', 'walkMeOptOut'];
    }
    async commandAction(logger) {
        const requestOptions = {
            url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/listtenantsettings?api-version=2020-10-01`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        try {
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PpTenantSettingsListCommand();
//# sourceMappingURL=tenant-settings-list.js.map