import { formatting } from './formatting.js';
import { cli } from '../cli/cli.js';
import { odata } from './odata.js';
const graphResource = 'https://graph.microsoft.com';
export const spe = {
    /**
     * Get the ID of a container type by its name.
     * @param name Name of the container type to search for
     * @returns ID of the container type
     */
    async getContainerTypeIdByName(name) {
        const containerTypes = await odata.getAllItems(`${graphResource}/beta/storage/fileStorage/containerTypes?$select=id,name&$filter=name eq '${formatting.encodeQueryParameter(name)}'`);
        if (containerTypes.length === 0) {
            throw new Error(`The specified container type '${name}' does not exist.`);
        }
        if (containerTypes.length > 1) {
            const containerTypeKeyValuePair = formatting.convertArrayToHashTable('id', containerTypes);
            const containerType = await cli.handleMultipleResultsFound(`Multiple container types with name '${name}' found.`, containerTypeKeyValuePair);
            return containerType.id;
        }
        return containerTypes[0].id;
    },
    /**
     * Get the ID of a container by its name.
     * @param containerTypeId ID of the container type.
     * @param name Name of the container to search for.
     * @returns ID of the container.
     */
    async getContainerIdByName(containerTypeId, name) {
        const containers = await odata.getAllItems(`${graphResource}/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`);
        const matchingContainers = containers.filter(c => c.displayName.toLowerCase() === name.toLowerCase());
        if (matchingContainers.length === 0) {
            throw new Error(`The specified container '${name}' does not exist.`);
        }
        if (matchingContainers.length > 1) {
            const containerKeyValuePair = formatting.convertArrayToHashTable('id', matchingContainers);
            const container = await cli.handleMultipleResultsFound(`Multiple containers with name '${name}' found.`, containerKeyValuePair);
            return container.id;
        }
        return matchingContainers[0].id;
    }
};
//# sourceMappingURL=spe.js.map