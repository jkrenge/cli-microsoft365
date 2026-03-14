import request from '../../../../request.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
export const options = z.strictObject({ ...globalOptionsZod.shape });
class PaEnvironmentListCommand extends PowerAppsCommand {
    get name() {
        return commands.ENVIRONMENT_LIST;
    }
    get description() {
        return 'Lists Microsoft Power Apps environments in the current tenant';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['name', 'displayName'];
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of Microsoft Power Apps environments...`);
        }
        const requestOptions = {
            url: `${this.resource}/providers/Microsoft.PowerApps/environments?api-version=2017-08-01`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            if (res.value.length > 0) {
                res.value.forEach(e => {
                    e.displayName = e.properties.displayName;
                });
            }
            await logger.log(res.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new PaEnvironmentListCommand();
//# sourceMappingURL=environment-list.js.map