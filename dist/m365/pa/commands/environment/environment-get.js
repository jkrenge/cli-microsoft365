import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    name: z.string().optional().alias('n'),
    default: z.boolean().optional()
});
class PaEnvironmentGetCommand extends PowerAppsCommand {
    get name() {
        return commands.ENVIRONMENT_GET;
    }
    get description() {
        return 'Gets information about the specified Microsoft Power Apps environment';
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
            await logger.logToStderr(`Retrieving information about Microsoft Power Apps environment ${args.options.name || 'default'}...`);
        }
        const environmentName = args.options.default ? '~default' : formatting.encodeQueryParameter(args.options.name);
        const requestOptions = {
            url: `${this.resource}/providers/Microsoft.PowerApps/environments/${environmentName}?api-version=2016-11-01`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            res.displayName = res.properties.displayName;
            res.provisioningState = res.properties.provisioningState;
            res.environmentSku = res.properties.environmentSku;
            res.azureRegionHint = res.properties.azureRegionHint;
            res.isDefault = res.properties.isDefault;
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PaEnvironmentGetCommand();
//# sourceMappingURL=environment-get.js.map