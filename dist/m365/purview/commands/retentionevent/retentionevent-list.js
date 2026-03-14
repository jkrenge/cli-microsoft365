import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({ ...globalOptionsZod.shape });
class PurviewRetentionEventListCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONEVENT_LIST;
    }
    get description() {
        return 'Get a list of retention events';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'displayName', 'eventTriggerDateTime'];
    }
    async commandAction(logger) {
        try {
            if (this.verbose) {
                await logger.logToStderr('Retrieving Purview retention events');
            }
            const items = await odata.getAllItems(`${this.resource}/v1.0/security/triggers/retentionEvents`);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PurviewRetentionEventListCommand();
//# sourceMappingURL=retentionevent-list.js.map