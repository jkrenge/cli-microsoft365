import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import { spe } from '../../../../utils/spe.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.string().optional().alias('i'),
    name: z.string().optional().alias('n'),
    containerTypeId: z.uuid().optional(),
    containerTypeName: z.string().optional(),
    recycle: z.boolean().optional(),
    force: z.boolean().optional().alias('f')
});
class SpeContainerRemoveCommand extends GraphCommand {
    get name() {
        return commands.CONTAINER_REMOVE;
    }
    get description() {
        return 'Removes a container';
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
            error: 'Options containerTypeId and containerTypeName are only required when deleting a container by name.'
        });
    }
    async commandAction(logger, args) {
        if (!args.options.force) {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove container '${args.options.id || args.options.name}'${!args.options.recycle ? ' permanently' : ''}?` });
            if (!result) {
                return;
            }
        }
        try {
            const containerId = await this.getContainerId(args.options, logger);
            if (this.verbose) {
                await logger.logToStderr(`Removing container with ID '${containerId}'...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/storage/fileStorage/containers/${containerId}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            if (args.options.recycle) {
                await request.delete(requestOptions);
                return;
            }
            // Container should be permanently deleted
            requestOptions.url += '/permanentDelete';
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getContainerId(options, logger) {
        if (options.id) {
            return options.id;
        }
        const containerTypeId = await this.getContainerTypeId(options, logger);
        if (this.verbose) {
            await logger.logToStderr(`Getting container ID for container with name '${options.name}'...`);
        }
        return spe.getContainerIdByName(containerTypeId, options.name);
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
export default new SpeContainerRemoveCommand();
//# sourceMappingURL=container-remove.js.map