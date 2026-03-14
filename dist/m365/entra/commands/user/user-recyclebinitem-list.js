import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = globalOptionsZod.strict();
class EntraUserRecycleBinItemListCommand extends GraphCommand {
    get name() {
        return commands.USER_RECYCLEBINITEM_LIST;
    }
    get description() {
        return 'Lists users from the recycle bin in the current tenant';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'displayName', 'userPrincipalName'];
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr('Retrieving users from the recycle bin...');
        }
        try {
            const users = await odata.getAllItems(`${this.resource}/v1.0/directory/deletedItems/microsoft.graph.user`);
            await logger.log(users);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraUserRecycleBinItemListCommand();
//# sourceMappingURL=user-recyclebinitem-list.js.map