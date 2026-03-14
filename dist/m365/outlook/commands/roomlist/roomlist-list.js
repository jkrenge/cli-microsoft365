import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({ ...globalOptionsZod.shape });
class OutlookRoomListListCommand extends GraphCommand {
    get name() {
        return commands.ROOMLIST_LIST;
    }
    get description() {
        return 'Get a collection of available roomlists';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'displayName', 'phone', 'emailAddress'];
    }
    async commandAction(logger) {
        try {
            const roomLists = await odata.getAllItems(`${this.resource}/v1.0/places/microsoft.graph.roomlist`);
            await logger.log(roomLists);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new OutlookRoomListListCommand();
//# sourceMappingURL=roomlist-list.js.map