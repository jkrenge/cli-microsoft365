import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    properties: z.string().optional().alias('p')
});
class EntraOrganizationListCommand extends GraphCommand {
    get name() {
        return commands.ORGANIZATION_LIST;
    }
    get description() {
        return 'Lists all Microsoft Entra ID organizations';
    }
    defaultProperties() {
        return ['id', 'displayName', 'tenantType'];
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            let url = `${this.resource}/v1.0/organization`;
            if (args.options.properties) {
                url += `?$select=${args.options.properties}`;
            }
            const requestOptions = {
                url: url,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            if (args.options.verbose) {
                await logger.logToStderr(`Retrieving organizations...`);
            }
            const res = await odata.getAllItems(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraOrganizationListCommand();
//# sourceMappingURL=organization-list.js.map