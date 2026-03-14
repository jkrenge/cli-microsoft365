import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    webUrl: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
        .alias('u'),
    id: z.int().positive()
});
class SpoNavigationNodeGetCommand extends SpoCommand {
    get name() {
        return commands.NAVIGATION_NODE_GET;
    }
    get description() {
        return 'Retrieve information about a specific navigation node';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about navigation node with id ${args.options.id}`);
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/navigation/GetNodeById(${args.options.id})?$expand=Children,Children/Children,Children/Children/Children`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const listInstance = await request.get(requestOptions);
            if (listInstance['odata.null']) {
                throw `No navigation node found with id ${args.options.id}.`;
            }
            await logger.log(listInstance);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoNavigationNodeGetCommand();
//# sourceMappingURL=navigation-node-get.js.map