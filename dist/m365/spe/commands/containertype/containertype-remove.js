import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import config from '../../../../config.js';
import { formatting } from '../../../../utils/formatting.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.uuid().optional().alias('i'),
    name: z.string().optional().alias('n'),
    force: z.boolean().optional().alias('f')
});
class SpeContainerTypeRemoveCommand extends SpoCommand {
    get name() {
        return commands.CONTAINERTYPE_REMOVE;
    }
    get description() {
        return 'Remove a specific container type';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.id, options.name].filter(o => o !== undefined).length === 1, {
            error: 'Use one of the following options: id, name.'
        });
    }
    async commandAction(logger, args) {
        if (!args.options.force) {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove container type ${args.options.id || args.options.name}?` });
            if (!result) {
                return;
            }
        }
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
            const containerTypeId = await this.getContainerTypeId(args.options, spoAdminUrl, logger);
            if (this.verbose) {
                await logger.logToStderr(`Removing container type '${args.options.id || args.options.name}'...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="7" ObjectPathId="6" /><Method Name="RemoveSPOContainerType" Id="8" ObjectPathId="6"><Parameters><Parameter TypeId="{b66ab1ca-fd51-44f9-8cfc-01f5c2a21f99}"><Property Name="ContainerTypeId" Type="Guid">{${containerTypeId}}</Property></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="6" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
            };
            const result = await request.post(requestOptions);
            if (result[0].ErrorInfo) {
                throw result[0].ErrorInfo.ErrorMessage;
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getContainerTypeId(options, spoAdminUrl, logger) {
        if (options.id) {
            return options.id;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving container type id for container type '${options.name}'...`);
        }
        const requestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="46" ObjectPathId="45" /><Method Name="GetSPOContainerTypes" Id="47" ObjectPathId="45"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="45" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };
        const json = await request.post(requestOptions);
        const response = json[0];
        if (response.ErrorInfo) {
            throw new Error(response.ErrorInfo.ErrorMessage);
        }
        const allContainerTypes = json[json.length - 1];
        const containerTypes = allContainerTypes.filter(ct => ct.DisplayName.toLowerCase() === options.name.toLowerCase());
        if (containerTypes.length === 0) {
            throw new Error(`The specified container type '${options.name}' does not exist.`);
        }
        if (containerTypes.length > 1) {
            const containerTypeKeyValuePair = formatting.convertArrayToHashTable('ContainerTypeId', containerTypes);
            const containerType = await cli.handleMultipleResultsFound(`Multiple container types with name '${options.name}' found.`, containerTypeKeyValuePair);
            return formatting.extractCsomGuid(containerType.ContainerTypeId);
        }
        return formatting.extractCsomGuid(containerTypes[0].ContainerTypeId);
    }
}
export default new SpeContainerTypeRemoveCommand();
//# sourceMappingURL=containertype-remove.js.map