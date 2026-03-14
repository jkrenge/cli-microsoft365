import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    name: z.string().optional().alias('n'),
    default: z.boolean().optional(),
    asAdmin: z.boolean().optional()
});
class PpEnvironmentGetCommand extends PowerPlatformCommand {
    get name() {
        return commands.ENVIRONMENT_GET;
    }
    get description() {
        return 'Gets information about the specified Power Platform environment';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => !!options.name !== !!options.default, {
            error: `Specify either name or default, but not both.`
        });
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving environment: ${args.options.name || 'default'}`);
        }
        let url = `${this.resource}/providers/Microsoft.BusinessAppPlatform`;
        if (args.options.asAdmin) {
            url += '/scopes/admin';
        }
        const envName = args.options.default ? '~Default' : formatting.encodeQueryParameter(args.options.name);
        url += `/environments/${envName}?api-version=2020-10-01`;
        const requestOptions = {
            url: url,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        await logger.log(response);
    }
}
export default new PpEnvironmentGetCommand();
//# sourceMappingURL=environment-get.js.map