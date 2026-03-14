import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { spe } from '../../../../utils/spe.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.string().optional().alias('i'),
    name: z.string().optional().alias('n'),
    containerTypeId: z.uuid().optional(),
    containerTypeName: z.string().optional(),
    force: z.boolean().optional().alias('f')
});
class SpeContainerRecycleBinItemRemoveCommand extends GraphCommand {
    get name() {
        return commands.CONTAINER_RECYCLEBINITEM_REMOVE;
    }
    get description() {
        return 'Permanently removes a container from the recycle bin';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine((options) => [options.id, options.name].filter(o => o !== undefined).length === 1, {
            error: 'Use one of the following options: id or name.'
        })
            .refine((options) => !options.name || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 1, {
            error: 'Use one of the following options when specifying the container name: containerTypeId or containerTypeName.'
        })
            .refine((options) => options.name || [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 0, {
            error: 'Options containerTypeId and containerTypeName are only required when removing a container by name.'
        });
    }
    async commandAction(logger, args) {
        if (!args.options.force) {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to permanently remove deleted container '${args.options.id || args.options.name}'?` });
            if (!result) {
                return;
            }
        }
        try {
            const containerId = await this.getContainerId(args.options, logger);
            if (this.verbose) {
                await logger.logToStderr(`Permanently removing deleted container with ID '${containerId}'...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/storage/fileStorage/deletedContainers/${containerId}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getContainerId(options, logger) {
        if (options.id) {
            return options.id;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving container with name '${options.name}'...`);
        }
        const containerTypeId = await this.getContainerTypeId(options, logger);
        const containers = await odata.getAllItems(`${this.resource}/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`);
        const matchingContainers = containers.filter(c => c.displayName.toLowerCase() === options.name.toLowerCase());
        if (matchingContainers.length === 0) {
            throw new Error(`The specified container '${options.name}' does not exist.`);
        }
        if (matchingContainers.length > 1) {
            const containerKeyValuePair = formatting.convertArrayToHashTable('id', matchingContainers);
            const container = await cli.handleMultipleResultsFound(`Multiple containers with name '${options.name}' found.`, containerKeyValuePair);
            return container.id;
        }
        return matchingContainers[0].id;
    }
    async getContainerTypeId(options, logger) {
        if (options.containerTypeId) {
            return options.containerTypeId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Getting container type with name '${options.containerTypeName}'...`);
        }
        return spe.getContainerTypeIdByName(options.containerTypeName);
    }
}
export default new SpeContainerRecycleBinItemRemoveCommand();
//# sourceMappingURL=container-recyclebinitem-remove.js.map