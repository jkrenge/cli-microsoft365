import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class VivaEngageCommunityListCommand extends GraphCommand {
    get name() {
        return commands.ENGAGE_COMMUNITY_LIST;
    }
    get description() {
        return 'Lists all Viva Engage communities';
    }
    defaultProperties() {
        return ['id', 'displayName', 'privacy'];
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr('Getting all Viva Engage communities...');
        }
        try {
            const results = await odata.getAllItems(`${this.resource}/v1.0/employeeExperience/communities`);
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new VivaEngageCommunityListCommand();
//# sourceMappingURL=engage-community-list.js.map