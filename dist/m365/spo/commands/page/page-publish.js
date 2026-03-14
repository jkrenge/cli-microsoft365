import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { Page } from './Page.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    webUrl: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
        .alias('u'),
    name: z.string().alias('n')
});
class SpoPagePublishCommand extends SpoCommand {
    get name() {
        return commands.PAGE_PUBLISH;
    }
    get description() {
        return 'Publishes a modern page';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            // Remove leading slashes from the page name (page can be nested in folders)
            let pageName = urlUtil.removeLeadingSlashes(args.options.name);
            if (!pageName.toLowerCase().endsWith('.aspx')) {
                pageName += '.aspx';
            }
            if (this.verbose) {
                await logger.logToStderr(`Publishing page ${pageName}...`);
            }
            await Page.publishPage(args.options.webUrl, pageName);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoPagePublishCommand();
//# sourceMappingURL=page-publish.js.map