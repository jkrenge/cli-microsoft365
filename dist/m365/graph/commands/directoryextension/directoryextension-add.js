import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { validation } from '../../../../utils/validation.js';
import { entraApp } from '../../../../utils/entraApp.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    name: z.string().alias('n'),
    appId: z.string().optional(),
    appObjectId: z.string().optional(),
    appName: z.string().optional(),
    dataType: z.enum(['Binary', 'Boolean', 'DateTime', 'Integer', 'LargeInteger', 'String']),
    targetObjects: z.string().transform((value) => value.split(',').map(String))
        .pipe(z.enum(['User', 'Group', 'Application', 'AdministrativeUnit', 'Device', 'Organization']).array()),
    isMultiValued: z.boolean().optional()
});
class GraphDirectoryExtensionAddCommand extends GraphCommand {
    get name() {
        return commands.DIRECTORYEXTENSION_ADD;
    }
    get description() {
        return 'Creates a new directory extension';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => Object.values([options.appId, options.appObjectId, options.appName]).filter(v => typeof v !== 'undefined').length === 1, {
            error: 'Specify either appId, appObjectId or appName, but not multiple'
        })
            .refine(options => (!options.appId && !options.appObjectId && !options.appName) || options.appObjectId || options.appName ||
            (options.appId && validation.isValidGuid(options.appId)), {
            error: e => `The '${e.input}' must be a valid GUID`,
            path: ['appId']
        })
            .refine(options => (!options.appId && !options.appObjectId && !options.appName) || options.appId || options.appName ||
            (options.appObjectId && validation.isValidGuid(options.appObjectId)), {
            error: e => `The '${e.input}' must be a valid GUID`,
            path: ['appObjectId']
        });
    }
    async commandAction(logger, args) {
        try {
            const appObjectId = await this.getAppObjectId(args.options);
            if (args.options.verbose) {
                await logger.logToStderr(`Adding directory extension to the app with id '${appObjectId}'...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/applications/${appObjectId}/extensionProperties`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                data: {
                    name: args.options.name,
                    dataType: args.options.dataType,
                    targetObjects: args.options.targetObjects,
                    isMultiValued: args.options.isMultiValued
                },
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppObjectId(options) {
        if (options.appObjectId) {
            return options.appObjectId;
        }
        if (options.appId) {
            return (await entraApp.getAppRegistrationByAppId(options.appId, ["id"])).id;
        }
        return (await entraApp.getAppRegistrationByAppName(options.appName, ["id"])).id;
    }
}
export default new GraphDirectoryExtensionAddCommand();
//# sourceMappingURL=directoryextension-add.js.map