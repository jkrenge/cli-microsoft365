import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    url: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
        .alias('u'),
    force: z.boolean().optional().alias('f')
});
class SpoHomeSiteRemoveCommand extends SpoCommand {
    get name() {
        return commands.HOMESITE_REMOVE;
    }
    get description() {
        return 'Removes a Home Site';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        const removeHomeSite = async () => {
            try {
                const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
                await this.removeHomeSiteByUrl(args.options.url, spoAdminUrl, logger);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeHomeSite();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove '${args.options.url}' as home site?` });
            if (result) {
                await removeHomeSite();
            }
        }
    }
    async removeHomeSiteByUrl(siteUrl, spoAdminUrl, logger) {
        const siteAdminProperties = await spo.getSiteAdminPropertiesByUrl(siteUrl, false, logger, this.verbose);
        if (this.verbose) {
            await logger.logToStderr(`Removing '${siteUrl}' as home site...`);
        }
        const requestOptions = {
            url: `${spoAdminUrl}/_api/SPO.Tenant/RemoveTargetedSite`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: {
                siteId: siteAdminProperties.SiteId
            }
        };
        await request.post(requestOptions);
    }
}
export default new SpoHomeSiteRemoveCommand();
//# sourceMappingURL=homesite-remove.js.map