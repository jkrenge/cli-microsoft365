import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    resourceId: z.string().alias('i'),
    resourceType: z.enum(['user', 'group', 'device', 'organization']).alias('t'),
    name: z.string().alias('n')
});
class GraphOpenExtensionGetCommand extends GraphCommand {
    get name() {
        return commands.OPENEXTENSION_GET;
    }
    get description() {
        return 'Retrieves a specific open extension for a resource';
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
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving open extension for resource ${args.options.resourceId}...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/${args.options.resourceType}${args.options.resourceType === 'organization' ? '' : 's'}/${args.options.resourceId}/extensions/${args.options.name}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new GraphOpenExtensionGetCommand();
//# sourceMappingURL=openextension-get.js.map