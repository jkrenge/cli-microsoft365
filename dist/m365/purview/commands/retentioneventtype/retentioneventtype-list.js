import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = globalOptionsZod.strict();
class PurviewRetentionEventTypeListCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONEVENTTYPE_LIST;
    }
    get description() {
        return 'Get a list of retention event types';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'displayName', 'createdDateTime'];
    }
    async commandAction(logger) {
        try {
            const items = await odata.getAllItems(`${this.resource}/v1.0/security/triggerTypes/retentionEventTypes`);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PurviewRetentionEventTypeListCommand();
//# sourceMappingURL=retentioneventtype-list.js.map