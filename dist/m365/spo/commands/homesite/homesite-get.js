import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    url: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
        .alias('u')
});
class SpoHomeSiteGetCommand extends SpoCommand {
    get name() {
        return commands.HOMESITE_GET;
    }
    get description() {
        return 'Gets information about a home site';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
            if (this.verbose) {
                await logger.log(`Retrieving home sites...`);
            }
            const homeSites = await odata.getAllItems(`${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`);
            const homeSite = homeSites.find(hs => urlUtil.removeTrailingSlashes(hs.Url).toLowerCase() === urlUtil.removeTrailingSlashes(args.options.url).toLowerCase());
            if (homeSite === undefined) {
                throw `Home site with URL '${args.options.url}' not found.`;
            }
            await logger.log(homeSite);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoHomeSiteGetCommand();
//# sourceMappingURL=homesite-get.js.map