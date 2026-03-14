import request from '../../../../request.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    asAdmin: z.boolean().optional()
});
class PpEnvironmentListCommand extends PowerPlatformCommand {
    get name() {
        return commands.ENVIRONMENT_LIST;
    }
    get description() {
        return 'Lists Microsoft Power Platform environments';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['name', 'displayName'];
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of Microsoft Power Platform environments...`);
        }
        let url;
        if (args.options.asAdmin) {
            url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments`;
        }
        else {
            url = `${this.resource}/providers/Microsoft.BusinessAppPlatform/environments`;
        }
        const requestOptions = {
            url: `${url}?api-version=2020-10-01`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            if (res.value && res.value.length > 0) {
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
export default new PpEnvironmentListCommand();
//# sourceMappingURL=environment-list.js.map