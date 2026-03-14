import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { spe } from '../../../../utils/spe.js';
import { odata } from '../../../../utils/odata.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    containerTypeId: z.uuid().optional(),
    containerTypeName: z.string().optional()
});
class SpeContainerRecycleBinItemListCommand extends GraphCommand {
    get name() {
        return commands.CONTAINER_RECYCLEBINITEM_LIST;
    }
    get description() {
        return 'Lists deleted containers of a specific container type';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['id', 'displayName'];
    }
    getRefinedSchema(schema) {
        return schema
            .refine((options) => [options.containerTypeId, options.containerTypeName].filter(o => o !== undefined).length === 1, {
            error: 'Use one of the following options: containerTypeId or containerTypeName.'
        });
    }
    async commandAction(logger, args) {
        try {
            const containerTypeId = await this.getContainerTypeId(args.options, logger);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving deleted containers of container type with ID '${containerTypeId}'...`);
            }
            const deletedContainers = await odata.getAllItems(`${this.resource}/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}`);
            await logger.log(deletedContainers);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getContainerTypeId(options, logger) {
        if (options.containerTypeId) {
            return options.containerTypeId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving container type id for container type '${options.containerTypeName}'...`);
        }
        return spe.getContainerTypeIdByName(options.containerTypeName);
    }
}
export default new SpeContainerRecycleBinItemListCommand();
//# sourceMappingURL=container-recyclebinitem-list.js.map