import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
const options = z.looseObject({
    ...globalOptionsZod.shape,
    name: z.string().alias('n'),
    resourceId: z.string().alias('i'),
    resourceType: z.enum(['user', 'group', 'device', 'organization']).alias('t'),
    keepUnchangedProperties: z.boolean().optional().alias('k')
});
class GraphOpenExtensionSetCommand extends GraphCommand {
    constructor() {
        super(...arguments);
        this.defaultOpenExtensionProperties = ['id', 'extensionName'];
    }
    get name() {
        return commands.OPENEXTENSION_SET;
    }
    get description() {
        return 'Updates an open extension for a resource';
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
            const currentExtension = await this.getOpenExtension(logger, args);
            const currentExtensionNames = Object.getOwnPropertyNames(currentExtension);
            const requestBody = {};
            requestBody["@odata.type"] = '#microsoft.graph.openTypeExtension';
            const unknownOptions = optionsUtils.getUnknownOptions(args.options, zod.schemaToOptions(this.schema));
            const unknownOptionsNames = Object.getOwnPropertyNames(unknownOptions);
            unknownOptionsNames.forEach(async (option) => {
                const value = unknownOptions[option];
                if (value === "") {
                    requestBody[option] = null;
                }
                else {
                    try {
                        const jsonObject = JSON.parse(value);
                        requestBody[option] = jsonObject;
                    }
                    catch {
                        requestBody[option] = value;
                    }
                }
            });
            currentExtensionNames.forEach(async (name) => {
                if (!unknownOptionsNames.includes(name) && (args.options.keepUnchangedProperties || this.defaultOpenExtensionProperties.includes(name))) {
                    requestBody[name] = currentExtension[name];
                }
            });
            const requestOptions = {
                url: `${this.resource}/v1.0/${args.options.resourceType}${args.options.resourceType === 'organization' ? '' : 's'}/${args.options.resourceId}/extensions/${args.options.name}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                data: requestBody,
                responseType: 'json'
            };
            if (args.options.verbose) {
                await logger.logToStderr(`Updating open extension of the ${args.options.resourceType} with id '${args.options.resourceId}'...`);
            }
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getOpenExtension(logger, args) {
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
        return await request.get(requestOptions);
    }
}
export default new GraphOpenExtensionSetCommand();
//# sourceMappingURL=openextension-set.js.map