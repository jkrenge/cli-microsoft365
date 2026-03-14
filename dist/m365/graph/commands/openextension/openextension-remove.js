import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    resourceId: z.string().alias('i'),
    resourceType: z.enum(['user', 'group', 'device', 'organization']).alias('t'),
    name: z.string().alias('n'),
    force: z.boolean().optional().alias('f')
});
class GraphOpenExtensionRemoveCommand extends GraphCommand {
    get name() {
        return commands.OPENEXTENSION_REMOVE;
    }
    get description() {
        return 'Removes a specific open extension for a resource';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => options.resourceType !== 'group' && options.resourceType !== 'device' && options.resourceType !== 'organization' || (options.resourceId && validation.isValidGuid(options.resourceId)), {
            error: e => `The '${e.input}' must be a valid GUID`,
            path: ['resourceId']
        })
            .refine(options => options.resourceType !== 'user' || (options.resourceId && (validation.isValidGuid(options.resourceId) || validation.isValidUserPrincipalName(options.resourceId))), {
            error: e => `The '${e.input}' must be a valid GUID or user principal name`,
            path: ['resourceId']
        });
    }
    async commandAction(logger, args) {
        const removeOpenExtension = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Removing open extension for resource ${args.options.resourceId}...`);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/${args.options.resourceType}${args.options.resourceType === 'organization' ? '' : 's'}/${args.options.resourceId}/extensions/${args.options.name}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    }
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeOpenExtension();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove open extension '${args.options.name}' from resource '${args.options.resourceId}' of type '${args.options.resourceType}'?` });
            if (result) {
                await removeOpenExtension();
            }
        }
    }
}
export default new GraphOpenExtensionRemoveCommand();
//# sourceMappingURL=openextension-remove.js.map