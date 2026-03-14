import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({ ...globalOptionsZod.shape });
class BookingBusinessListCommand extends GraphCommand {
    get name() {
        return commands.BUSINESS_LIST;
    }
    get description() {
        return 'Lists all Microsoft Bookings businesses that are created for the tenant.';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'displayName'];
    }
    async commandAction(logger) {
        const endpoint = `${this.resource}/v1.0/solutions/bookingBusinesses`;
        try {
            const items = await odata.getAllItems(endpoint);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new BookingBusinessListCommand();
//# sourceMappingURL=business-list.js.map