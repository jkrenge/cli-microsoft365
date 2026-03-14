import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { validation } from '../../../../utils/validation.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.uuid().optional().alias('i'),
    displayName: z.string().optional().alias('n'),
    properties: z.string().optional().alias('p')
});
class EntraRoleDefinitionGetCommand extends GraphCommand {
    get name() {
        return commands.ROLEDEFINITION_GET;
    }
    get description() {
        return 'Gets a specific Microsoft Entra ID role definition';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => !options.id !== !options.displayName, {
            error: 'Specify either id or displayName, but not both'
        })
            .refine(options => options.id || options.displayName, {
            error: 'Specify either id or displayName'
        })
            .refine(options => (!options.id && !options.displayName) || options.displayName || (options.id && validation.isValidGuid(options.id)), {
            error: e => `The '${e.input}' must be a valid GUID`,
            path: ['id']
        });
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Getting Microsoft Entra ID role definition...');
        }
        try {
            let result;
            if (args.options.id) {
                result = await roleDefinition.getRoleDefinitionById(args.options.id, args.options.properties);
            }
            else {
                result = await roleDefinition.getRoleDefinitionByDisplayName(args.options.displayName, args.options.properties);
            }
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraRoleDefinitionGetCommand();
//# sourceMappingURL=roledefinition-get.js.map