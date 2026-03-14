import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
const options = z.looseObject({
    ...globalOptionsZod.shape,
    name: z.string().alias('n'),
    resourceId: z.string().alias('i'),
    resourceType: z.enum(['user', 'group', 'device', 'organization']).alias('t')
});
class GraphOpenExtensionAddCommand extends GraphCommand {
    get name() {
        return commands.OPENEXTENSION_ADD;
    }
    get description() {
        return 'Adds an open extension to a resource';
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
            const requestBody = {};
            requestBody["extensionName"] = args.options.name;
            const unknownOptions = optionsUtils.getUnknownOptions(args.options, zod.schemaToOptions(this.schema));
            const unknownOptionsNames = Object.getOwnPropertyNames(unknownOptions);
            unknownOptionsNames.forEach(async (o) => {
                try {
                    const jsonObject = JSON.parse(unknownOptions[o]);
                    requestBody[o] = jsonObject;
                }
                catch {
                    requestBody[o] = unknownOptions[o];
                }
            });
            const requestOptions = {
                url: `${this.resource}/v1.0/${args.options.resourceType}${args.options.resourceType === 'organization' ? '' : 's'}/${args.options.resourceId}/extensions`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                data: requestBody,
                responseType: 'json'
            };
            if (args.options.verbose) {
                await logger.logToStderr(`Adding open extension to the ${args.options.resourceType} with id '${args.options.resourceId}'...`);
            }
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new GraphOpenExtensionAddCommand();
//# sourceMappingURL=openextension-add.js.map