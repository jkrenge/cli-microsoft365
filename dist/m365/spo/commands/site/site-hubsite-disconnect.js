import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { spo } from '../../../../utils/spo.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    siteUrl: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
        .alias('u'),
    asAdmin: z.boolean().optional(),
    force: z.boolean().alias('f').optional()
});
class SpoSiteHubSiteDisconnectCommand extends SpoCommand {
    get name() {
        return commands.SITE_HUBSITE_DISCONNECT;
    }
    get description() {
        return 'Disconnects the specified site collection from its hub site';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.disconnectHubSite(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to disconnect the site collection ${args.options.siteUrl} from its hub site?` });
            if (result) {
                await this.disconnectHubSite(logger, args);
            }
        }
    }
    async disconnectHubSite(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Disconnecting site collection ${args.options.siteUrl} from its hub site...`);
            }
            const requestOptions = {
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            if (!args.options.asAdmin) {
                requestOptions.url = `${args.options.siteUrl}/_api/site/JoinHubSite('00000000-0000-0000-0000-000000000000')`;
            }
            else {
                const tenantAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
                requestOptions.url = `${tenantAdminUrl}/_api/SPO.Tenant/ConnectSiteToHubSiteById`;
                requestOptions.data = {
                    siteUrl: args.options.siteUrl,
                    hubSiteId: '00000000-0000-0000-0000-000000000000'
                };
            }
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoSiteHubSiteDisconnectCommand();
//# sourceMappingURL=site-hubsite-disconnect.js.map