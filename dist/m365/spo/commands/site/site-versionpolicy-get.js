import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    siteUrl: z.string().alias('u')
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
});
class SpoSiteVersionpolicyGetCommand extends SpoCommand {
    get name() {
        return commands.SITE_VERSIONPOLICY_GET;
    }
    get description() {
        return 'Retrieves the version policy settings of a specific site';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving version policy settings for site '${args.options.siteUrl}'...`);
        }
        try {
            const requestOptions = {
                url: `${args.options.siteUrl}/_api/site/VersionPolicyForNewLibrariesTemplate?$expand=VersionPolicies`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const response = await request.get(requestOptions);
            let defaultTrimMode = 'number';
            if (response.MajorVersionLimit === -1) {
                defaultTrimMode = 'inheritTenant';
            }
            else if (response.VersionPolicies) {
                switch (response.VersionPolicies.DefaultTrimMode) {
                    case 1:
                        defaultTrimMode = 'age';
                        break;
                    case 2:
                        defaultTrimMode = 'automatic';
                        break;
                    case 0:
                    default:
                        defaultTrimMode = 'number';
                }
            }
            const output = {
                defaultTrimMode: defaultTrimMode,
                defaultExpireAfterDays: response.VersionPolicies?.DefaultExpireAfterDays ?? null,
                majorVersionLimit: response.MajorVersionLimit
            };
            await logger.log(output);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoSiteVersionpolicyGetCommand();
//# sourceMappingURL=site-versionpolicy-get.js.map