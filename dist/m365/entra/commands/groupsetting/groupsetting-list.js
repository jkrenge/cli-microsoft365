import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupSettingListCommand extends GraphCommand {
    get name() {
        return commands.GROUPSETTING_LIST;
    }
    get description() {
        return 'Lists Entra group settings';
    }
    defaultProperties() {
        return ['id', 'displayName'];
    }
    async commandAction(logger) {
        try {
            const groupSettings = await odata.getAllItems(`${this.resource}/v1.0/groupSettings`);
            await logger.log(groupSettings);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraGroupSettingListCommand();
//# sourceMappingURL=groupsetting-list.js.map