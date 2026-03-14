import { cli } from '../cli/cli.js';
import { formatting } from './formatting.js';
import { odata } from './odata.js';
export const customAppScope = {
    /**
     * Get a custom application scope by its name
     * @param displayName Custom application scope display name.
     * @param properties Comma-separated list of properties to include in the response.
     * @returns The custom application scope.
     * @throws Error when role definition was not found.
     */
    async getCustomAppScopeByDisplayName(displayName, properties) {
        let url = `https://graph.microsoft.com/beta/roleManagement/exchange/customAppScopes?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`;
        if (properties) {
            url += `&$select=${properties}`;
        }
        const customAppScopes = await odata.getAllItems(url);
        if (customAppScopes.length === 0) {
            throw new Error(`The specified custom application scope '${displayName}' does not exist.`);
        }
        if (customAppScopes.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', customAppScopes);
            const selectedCustomAppScope = await cli.handleMultipleResultsFound(`Multiple custom application scopes with name '${displayName}' found.`, resultAsKeyValuePair);
            return selectedCustomAppScope;
        }
        return customAppScopes[0];
    }
};
//# sourceMappingURL=customAppScope.js.map