import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    resourceId: z.string().alias('i'),
    resourceType: z.enum(['user', 'group', 'device', 'organization']).alias('t')
});
class GraphOpenExtensionListCommand extends GraphCommand {
    get name() {
        return commands.OPENEXTENSION_LIST;
    }
    get description() {
        return 'Retrieves all open extensions for a resource';
    }
    defaultProperties() {
        return ['id', 'extensionName'];
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => options.resourceType !== 'group' && options.resourceType !== 'device' && options.resourceType !== 'organization' ||
            (options.resourceId && validation.isValidGuid(options.resourceId)), {
            error: e => `The '${e.input}' must be a valid GUID`,
            path: ['resourceId']
        })
            .refine(options => options.resourceType !== 'user' ||
            (options.resourceId && (validation.isValidGuid(options.resourceId) || validation.isValidUserPrincipalName(options.resourceId))), {
            error: e => `The '${e.input}' must be a valid GUID or user principal name`,
            path: ['resourceId']
        });
    }
    async commandAction(logger, args) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/${args.options.resourceType}${args.options.resourceType === 'organization' ? '' : 's'}/${args.options.resourceId}/extensions`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            if (args.options.verbose) {
                await logger.logToStderr(`Retrieving open extensions for the ${args.options.resourceType} with id '${args.options.resourceId}'...`);
            }
            const res = await odata.getAllItems(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new GraphOpenExtensionListCommand();
//# sourceMappingURL=openextension-list.js.map