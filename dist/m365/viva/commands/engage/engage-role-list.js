import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class VivaEngageRoleListCommand extends GraphCommand {
    get name() {
        return commands.ENGAGE_ROLE_LIST;
    }
    get description() {
        return 'Lists all Viva Engage roles';
    }
    defaultProperties() {
        return ['id', 'displayName'];
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr('Getting all Viva Engage roles...');
        }
        try {
            const results = await odata.getAllItems(`${this.resource}/beta/employeeExperience/roles`);
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new VivaEngageRoleListCommand();
//# sourceMappingURL=engage-role-list.js.map