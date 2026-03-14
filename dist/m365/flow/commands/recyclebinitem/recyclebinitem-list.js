import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { cli } from '../../../../cli/cli.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    environmentName: z.string().alias('e')
});
class FlowRecycleBinItemListCommand extends PowerAutomateCommand {
    get name() {
        return commands.RECYCLEBINITEM_LIST;
    }
    get description() {
        return 'Lists all soft-deleted Power Automate flows within an environment';
    }
    defaultProperties() {
        return ['name', 'displayName'];
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Getting list of soft-deleted flows in environment ${args.options.environmentName}...`);
            }
            const flows = await odata.getAllItems(`${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/v2/flows?api-version=2016-11-01&include=softDeletedFlows`);
            const deletedFlows = flows.filter(flow => flow.properties.state === 'Deleted');
            if (cli.shouldTrimOutput(args.options.output)) {
                deletedFlows.forEach(flow => {
                    flow.displayName = flow.properties.displayName;
                });
            }
            await logger.log(deletedFlows);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new FlowRecycleBinItemListCommand();
//# sourceMappingURL=recyclebinitem-list.js.map