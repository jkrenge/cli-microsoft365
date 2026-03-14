import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    url: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
        error: 'Specify a valid SharePoint site URL'
    }).alias('u')
});
class SpoSiteGetCommand extends SpoCommand {
    get name() {
        return commands.SITE_GET;
    }
    get description() {
        return 'Gets information about the specific site collection';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${args.options.url}/_api/site`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoSiteGetCommand();
//# sourceMappingURL=site-get.js.map