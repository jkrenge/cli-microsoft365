import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { spe } from '../../../../utils/spe.js';
import GraphCommand from '../../../base/GraphCommand.js';
import request from '../../../../request.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    name: z.string().alias('n'),
    description: z.string().optional().alias('d'),
    containerTypeId: z.uuid().optional(),
    containerTypeName: z.string().optional(),
    ocrEnabled: z.boolean().optional(),
    itemMajorVersionLimit: z.number()
        .refine(numb => validation.isValidPositiveInteger(numb), {
        error: e => `'${e.input}' is not a valid positive integer.`
    }).optional(),
    itemVersioningEnabled: z.boolean().optional()
});
class SpeContainerAddCommand extends GraphCommand {
    get name() {
        return commands.CONTAINER_ADD;
    }
    get description() {
        return 'Creates a new container';
    }
    get schema() {
        return options;
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
                await logger.logToStderr(`Creating container with name '${args.options.name}'...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/storage/fileStorage/containers`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    displayName: args.options.name,
                    description: args.options.description,
                    containerTypeId: containerTypeId,
                    settings: {
                        isOcrEnabled: args.options.ocrEnabled,
                        itemMajorVersionLimit: args.options.itemMajorVersionLimit,
                        isItemVersioningEnabled: args.options.itemVersioningEnabled
                    }
                }
            };
            const container = await request.post(requestOptions);
            await logger.log(container);
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
            await logger.logToStderr(`Getting container type with name '${options.containerTypeName}'...`);
        }
        return spe.getContainerTypeIdByName(options.containerTypeName);
    }
}
export default new SpeContainerAddCommand();
//# sourceMappingURL=container-add.js.map