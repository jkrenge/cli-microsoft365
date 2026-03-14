import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
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
    id: z.uuid().alias('i'),
    asAdmin: z.boolean().optional()
});
class SpoSiteHubSiteConnectCommand extends SpoCommand {
    get name() {
        return commands.SITE_HUBSITE_CONNECT;
    }
    get description() {
        return 'Connects the specified site collection to the given hub site';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Connecting site collection ${args.options.siteUrl} to hub site ${args.options.id}...`);
            }
            const requestOptions = {
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            if (!args.options.asAdmin) {
                requestOptions.url = `${args.options.siteUrl}/_api/site/JoinHubSite('${formatting.encodeQueryParameter(args.options.id)}')`;
            }
            else {
                const tenantAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
                requestOptions.url = `${tenantAdminUrl}/_api/SPO.Tenant/ConnectSiteToHubSiteById`;
                requestOptions.data = {
                    siteUrl: args.options.siteUrl,
                    hubSiteId: args.options.id
                };
            }
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoSiteHubSiteConnectCommand();
//# sourceMappingURL=site-hubsite-connect.js.map