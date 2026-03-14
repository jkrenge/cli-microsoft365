import { odata } from '../../../../utils/odata.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
export const options = z.strictObject({ ...globalOptionsZod.shape });
class FlowEnvironmentListCommand extends PowerAutomateCommand {
    get name() {
        return commands.ENVIRONMENT_LIST;
    }
    get description() {
        return 'Lists Microsoft Flow environments in the current tenant';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['name', 'displayName'];
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of Microsoft Flow environments...`);
        }
        try {
            const res = await odata.getAllItems(`${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments?api-version=2016-11-01`);
            if (res.length > 0) {
                if (args.options.output !== 'json') {
                    res.forEach(e => {
                        e.displayName = e.properties.displayName;
                    });
                }
            }
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new FlowEnvironmentListCommand();
//# sourceMappingURL=environment-list.js.map