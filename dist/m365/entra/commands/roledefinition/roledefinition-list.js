import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    properties: z.string().optional().alias('p'),
    filter: z.string().optional().alias('f')
});
class EntraRoleDefinitionListCommand extends GraphCommand {
    get name() {
        return commands.ROLEDEFINITION_LIST;
    }
    get description() {
        return 'Lists all Microsoft Entra ID role definitions';
    }
    defaultProperties() {
        return ['id', 'displayName', 'isBuiltIn', 'isEnabled'];
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Getting Microsoft Entra ID role definitions...');
        }
        try {
            const queryParameters = [];
            if (args.options.properties) {
                queryParameters.push(`$select=${args.options.properties}`);
            }
            if (args.options.filter) {
                queryParameters.push(`$filter=${args.options.filter}`);
            }
            const queryString = queryParameters.length > 0
                ? `?${queryParameters.join('&')}`
                : '';
            const results = await odata.getAllItems(`${this.resource}/v1.0/roleManagement/directory/roleDefinitions${queryString}`);
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraRoleDefinitionListCommand();
//# sourceMappingURL=roledefinition-list.js.map