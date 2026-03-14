import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    name: z.string().optional().alias('n'),
    default: z.boolean().optional()
});
class FlowEnvironmentGetCommand extends PowerAutomateCommand {
    get name() {
        return commands.ENVIRONMENT_GET;
    }
    get description() {
        return 'Gets information about the specified Microsoft Flow environment';
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
            await logger.logToStderr(`Retrieving information about Microsoft Flow environment ${args.options.name ?? 'default'}...`);
        }
        let requestUrl = `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/`;
        requestUrl += args.options.default ? '~default' : formatting.encodeQueryParameter(args.options.name);
        const requestOptions = {
            url: `${requestUrl}?api-version=2016-11-01`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        try {
            const flowItem = await request.get(requestOptions);
            if (args.options.output !== 'json') {
                flowItem.displayName = flowItem.properties.displayName;
                flowItem.provisioningState = flowItem.properties.provisioningState;
                flowItem.environmentSku = flowItem.properties.environmentSku;
                flowItem.azureRegionHint = flowItem.properties.azureRegionHint;
                flowItem.isDefault = flowItem.properties.isDefault;
            }
            await logger.log(flowItem);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new FlowEnvironmentGetCommand();
//# sourceMappingURL=environment-get.js.map