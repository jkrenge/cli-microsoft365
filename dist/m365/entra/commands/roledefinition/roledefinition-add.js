import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    displayName: z.string().alias('n'),
    allowedResourceActions: z.string().transform((value) => value.split(',').map(String)).alias('a'),
    description: z.string().optional().alias('d'),
    enabled: z.boolean().optional().alias('e'),
    version: z.string().optional().alias('v')
});
class EntraRoleDefinitionAddCommand extends GraphCommand {
    get name() {
        return commands.ROLEDEFINITION_ADD;
    }
    get description() {
        return 'Creates a custom Microsoft Entra ID role definition';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (args.options.verbose) {
            await logger.logToStderr(`Creating custom role definition with name ${args.options.displayName}...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/roleManagement/directory/roleDefinitions`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            data: {
                displayName: args.options.displayName,
                rolePermissions: [
                    {
                        allowedResourceActions: args.options.allowedResourceActions
                    }
                ],
                description: args.options.description,
                isEnabled: args.options.enabled !== undefined ? args.options.enabled : true,
                version: args.options.version
            },
            responseType: 'json'
        };
        try {
            const result = await request.post(requestOptions);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraRoleDefinitionAddCommand();
//# sourceMappingURL=roledefinition-add.js.map