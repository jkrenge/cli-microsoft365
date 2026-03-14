import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class ExternalConnectionListCommand extends GraphCommand {
    get name() {
        return commands.CONNECTION_LIST;
    }
    get description() {
        return 'Lists external connections defined in the Microsoft Search';
    }
    alias() {
        return [commands.EXTERNALCONNECTION_LIST];
    }
    defaultProperties() {
        return ['id', 'name', 'state'];
    }
    async commandAction(logger) {
        try {
            const connections = await odata.getAllItems(`${this.resource}/v1.0/external/connections`);
            await logger.log(connections);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new ExternalConnectionListCommand();
//# sourceMappingURL=connection-list.js.map