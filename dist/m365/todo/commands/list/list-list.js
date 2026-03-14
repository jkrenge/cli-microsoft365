import { odata } from '../../../../utils/odata.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TodoListListCommand extends GraphDelegatedCommand {
    get name() {
        return commands.LIST_LIST;
    }
    get description() {
        return 'Returns a list of Microsoft To Do task lists';
    }
    defaultProperties() {
        return ['displayName', 'id'];
    }
    async commandAction(logger) {
        try {
            const items = await odata.getAllItems(`${this.resource}/v1.0/me/todo/lists`);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new TodoListListCommand();
//# sourceMappingURL=list-list.js.map