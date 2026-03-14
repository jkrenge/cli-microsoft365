import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class SpoServicePrincipalGrantListCommand extends GraphCommand {
    constructor() {
        super(...arguments);
        this.spoServicePrincipalDisplayName = 'SharePoint Online Web Client Extensibility';
    }
    get name() {
        return commands.SERVICEPRINCIPAL_GRANT_LIST;
    }
    get description() {
        return 'Lists permissions granted to the service principal';
    }
    alias() {
        return [commands.SP_GRANT_LIST];
    }
    async commandAction(logger) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving permissions granted to the service principal '${this.spoServicePrincipalDisplayName}'...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/servicePrincipals?$filter=displayName eq '${this.spoServicePrincipalDisplayName}'&$select=id`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const response = await request.get(requestOptions);
            if (response.value.length === 0) {
                throw `Service principal '${this.spoServicePrincipalDisplayName}' not found`;
            }
            requestOptions.url = `${this.resource}/v1.0/servicePrincipals/${response.value[0].id}/oauth2PermissionGrants`;
            const result = await request.get(requestOptions);
            await logger.log(result.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoServicePrincipalGrantListCommand();
//# sourceMappingURL=serviceprincipal-grant-list.js.map