import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { odata } from '../../../../utils/odata.js';
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
    location: z.enum(['QuickLaunch', 'TopNavigationBar']).alias('l')
});
class SpoNavigationNodeListCommand extends SpoCommand {
    get name() {
        return commands.NAVIGATION_NODE_LIST;
    }
    get description() {
        return 'Lists nodes from the specified site navigation';
    }
    defaultProperties() {
        return ['Id', 'Title', 'Url'];
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving navigation nodes...`);
        }
        try {
            const res = await odata.getAllItems(`${args.options.webUrl}/_api/web/navigation/${args.options.location.toLowerCase()}?$expand=Children,Children/Children,Children/Children/Children`);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoNavigationNodeListCommand();
//# sourceMappingURL=navigation-node-list.js.map