import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import PlannerCommand from '../../../base/PlannerCommand.js';
import commands from '../../commands.js';
export const options = globalOptionsZod.strict();
class PlannerTenantSettingsListCommand extends PlannerCommand {
    get name() {
        return commands.TENANT_SETTINGS_LIST;
    }
    get description() {
        return 'Lists the Microsoft Planner configuration of the tenant';
    }
    get schema() {
        return options;
    }
    async commandAction(logger) {
        const requestOptions = {
            url: `${this.resource}/taskAPI/tenantAdminSettings/Settings`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const result = await request.get(requestOptions);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PlannerTenantSettingsListCommand();
//# sourceMappingURL=tenant-settings-list.js.map