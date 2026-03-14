import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
import request from '../../../../request.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    environmentName: z.string().alias('e'),
    flowName: z.uuid().alias('n')
});
class FlowRecycleBinItemRestoreCommand extends PowerAutomateCommand {
    get name() {
        return commands.RECYCLEBINITEM_RESTORE;
    }
    get description() {
        return 'Restores a soft-deleted Power Automate flow';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Restoring soft-deleted flow ${args.options.flowName} from environment ${args.options.environmentName}...`);
            }
            const requestOptions = {
                url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${args.options.flowName}/restore?api-version=2016-11-01`,
                headers: {
                    accept: 'application/json'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new FlowRecycleBinItemRestoreCommand();
//# sourceMappingURL=recyclebinitem-restore.js.map