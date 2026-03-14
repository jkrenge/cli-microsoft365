import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { directoryExtension } from '../../../../utils/directoryExtension.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.uuid().optional().alias('i'),
    name: z.string().optional().alias('n'),
    appId: z.uuid().optional(),
    appObjectId: z.uuid().optional(),
    appName: z.string().optional()
});
class GraphDirectoryExtensionGetCommand extends GraphCommand {
    get name() {
        return commands.DIRECTORYEXTENSION_GET;
    }
    get description() {
        return 'Retrieves the definition of a directory extension';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => !options.id !== !options.name, {
            error: 'Specify either id or name, but not both'
        })
            .refine(options => options.id || options.name, {
            error: 'Specify either id or name'
        })
            .refine(options => Object.values([options.appId, options.appObjectId, options.appName]).filter(v => typeof v !== 'undefined').length === 1, {
            error: 'Specify either appId, appObjectId or appName, but not multiple'
        });
    }
    async commandAction(logger, args) {
        try {
            const appObjectId = await this.getAppObjectId(args.options);
            let schemeExtensionId = args.options.id;
            if (args.options.name) {
                schemeExtensionId = (await directoryExtension.getDirectoryExtensionByName(args.options.name, appObjectId, ['id'])).id;
            }
            if (args.options.verbose) {
                await logger.logToStderr(`Retrieving schema extension with ID ${schemeExtensionId} from application with ID ${appObjectId}...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/applications/${appObjectId}/extensionProperties/${schemeExtensionId}`,
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
export default new GraphDirectoryExtensionGetCommand();
//# sourceMappingURL=directoryextension-get.js.map