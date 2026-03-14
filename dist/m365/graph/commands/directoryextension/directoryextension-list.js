import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { odata } from '../../../../utils/odata.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    appId: z.uuid().optional(),
    appObjectId: z.uuid().optional(),
    appName: z.string().optional()
});
class GraphDirectoryExtensionListCommand extends GraphCommand {
    get name() {
        return commands.DIRECTORYEXTENSION_LIST;
    }
    get description() {
        return 'Retrieves a list of directory extensions';
    }
    defaultProperties() {
        return ['id', 'name', 'appDisplayName'];
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => ([options.appId, options.appObjectId, options.appName].filter(x => x !== undefined).length <= 1), {
            error: 'Specify either appId, appObjectId, or appName, but not multiple.'
        });
    }
    async commandAction(logger, args) {
        try {
            if (args.options.appId || args.options.appObjectId || args.options.appName) {
                const appObjectId = await this.getAppObjectId(args.options);
                const endpoint = `${this.resource}/v1.0/applications/${appObjectId}/extensionProperties/`;
                const items = await odata.getAllItems(endpoint);
                await logger.log(items);
            }
            else {
                const requestOptions = {
                    url: `${this.resource}/v1.0/directoryObjects/getAvailableExtensionProperties`,
                    headers: {
                        'content-type': 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                const res = await request.post(requestOptions);
                await logger.log(res.value);
            }
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
export default new GraphDirectoryExtensionListCommand();
//# sourceMappingURL=directoryextension-list.js.map