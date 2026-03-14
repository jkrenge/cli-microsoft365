import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    resourceNamespace: z.string().alias('n'),
    privileged: z.boolean().optional().alias('p')
});
class EntraRolePermissionListCommand extends GraphCommand {
    get name() {
        return commands.ROLEPERMISSION_LIST;
    }
    get description() {
        return 'Lists all Microsoft Entra ID role permissions';
    }
    defaultProperties() {
        return ['id', 'name', 'actionVerb', 'isPrivileged'];
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Getting Microsoft Entra ID role permissions...');
        }
        try {
            const queryString = args.options.privileged ? '?$filter=isPrivileged eq true' : '';
            const url = `${this.resource}/beta/roleManagement/directory/resourceNamespaces/${args.options.resourceNamespace}/resourceActions${queryString}`;
            const results = await odata.getAllItems(url);
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraRolePermissionListCommand();
//# sourceMappingURL=rolepermission-list.js.map