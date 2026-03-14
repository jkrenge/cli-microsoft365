import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraResourcenamespaceListCommand extends GraphCommand {
    get name() {
        return commands.RESOURCENAMESPACE_LIST;
    }
    get description() {
        return 'Get a list of the RBAC resource namespaces and their properties';
    }
    defaultProperties() {
        return ['id', 'name'];
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr('Getting a list of the RBAC resource namespaces and their properties...');
        }
        try {
            const results = await odata.getAllItems(`${this.resource}/beta/roleManagement/directory/resourceNamespaces`);
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraResourcenamespaceListCommand();
//# sourceMappingURL=resourcenamespace-list.js.map