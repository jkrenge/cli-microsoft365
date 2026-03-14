import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
export const options = globalOptionsZod.strict();
class EntraLicenseListCommand extends GraphCommand {
    get name() {
        return commands.LICENSE_LIST;
    }
    get description() {
        return 'Lists commercial subscriptions that an organization has acquired';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'skuId', 'skuPartNumber'];
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving the commercial subscriptions that an organization has acquired`);
        }
        try {
            const items = await odata.getAllItems(`${this.resource}/v1.0/subscribedSkus`);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraLicenseListCommand();
//# sourceMappingURL=license-list.js.map