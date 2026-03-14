import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import PowerBICommand from '../../../base/PowerBICommand.js';
import commands from '../../commands.js';
export const options = globalOptionsZod.strict();
class PpGatewayListCommand extends PowerBICommand {
    get name() {
        return commands.GATEWAY_LIST;
    }
    get description() {
        return 'Returns a list of gateways for which the user is an admin';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'name'];
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of gateways for which the user is an admin...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/myorg/gateways`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PpGatewayListCommand();
//# sourceMappingURL=gateway-list.js.map