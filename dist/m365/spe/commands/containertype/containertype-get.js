import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { formatting } from '../../../../utils/formatting.js';
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { cli } from '../../../../cli/cli.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    id: z.uuid().optional().alias('i'),
    name: z.string().optional().alias('n')
});
class SpeContainerTypeGetCommand extends GraphDelegatedCommand {
    get name() {
        return commands.CONTAINERTYPE_GET;
    }
    get description() {
        return 'Gets a container type';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.id, options.name].filter(o => o !== undefined).length === 1, {
            error: 'Use one of the following options: id or name.'
        });
    }
    async commandAction(logger, args) {
        try {
            let result;
            if (args.options.name) {
                const containerTypes = await odata.getAllItems(`${this.resource}/beta/storage/fileStorage/containerTypes?$filter=name eq '${formatting.encodeQueryParameter(args.options.name)}'`);
                if (containerTypes.length === 0) {
                    throw `The specified container type '${args.options.name}' does not exist.`;
                }
                if (containerTypes.length > 1) {
                    const containerKeyValuePair = formatting.convertArrayToHashTable('id', containerTypes);
                    result = await cli.handleMultipleResultsFound(`Multiple container types with name '${args.options.name}' found.`, containerKeyValuePair);
                }
                else {
                    result = containerTypes[0];
                }
            }
            else {
                const requestOptions = {
                    url: `${this.resource}/beta/storage/fileStorage/containerTypes/${args.options.id}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                result = await request.get(requestOptions);
            }
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpeContainerTypeGetCommand();
//# sourceMappingURL=containertype-get.js.map