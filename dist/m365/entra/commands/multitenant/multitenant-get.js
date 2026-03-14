import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraMultitenantGetCommand extends GraphCommand {
    get name() {
        return commands.MULTITENANT_GET;
    }
    get description() {
        return 'Gets properties of the multitenant organization';
    }
    async commandAction(logger) {
        const requestOptions = {
            url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            if (this.verbose) {
                await logger.logToStderr('Retrieving multitenant organization...');
            }
            const multitenantOrg = await request.get(requestOptions);
            await logger.log(multitenantOrg);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new EntraMultitenantGetCommand();
//# sourceMappingURL=multitenant-get.js.map