import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
export const options = globalOptionsZod
    .extend({
    displayName: z.string().alias('n'),
    description: z.string().optional().alias('d')
}).strict();
class EntraMultitenantAddCommand extends GraphCommand {
    get name() {
        return commands.MULTITENANT_ADD;
    }
    get description() {
        return 'Creates a new multitenant organization';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Creating multitenant organization...');
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                description: args.options.description,
                displayName: args.options.displayName
            }
        };
        try {
            const multitenantOrg = await request.put(requestOptions);
            await logger.log(multitenantOrg);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraMultitenantAddCommand();
//# sourceMappingURL=multitenant-add.js.map