import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = globalOptionsZod.strict();
class PurviewRetentionLabelListCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONLABEL_LIST;
    }
    get description() {
        return 'Get a list of retention labels';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'displayName', 'isInUse'];
    }
    async commandAction(logger) {
        try {
            const items = await odata.getAllItems(`${this.resource}/beta/security/labels/retentionLabels`);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PurviewRetentionLabelListCommand();
//# sourceMappingURL=retentionlabel-list.js.map