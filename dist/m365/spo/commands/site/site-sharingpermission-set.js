import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request from '../../../../request.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    siteUrl: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    }).alias('u'),
    capability: z.enum(['full', 'limited', 'ownersOnly'])
});
class SpoSiteSharingPermissionSetCommand extends SpoCommand {
    get name() {
        return commands.SITE_SHARINGPERMISSION_SET;
    }
    get description() {
        return 'Controls how a site and its components can be shared';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Updating sharing permissions for site '${args.options.siteUrl}'...`);
            }
            const { capability } = args.options;
            if (this.verbose) {
                await logger.logToStderr(`Updating site sharing permissions...`);
            }
            const requestOptionsWeb = {
                url: `${args.options.siteUrl}/_api/Web`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: {
                    MembersCanShare: capability === 'full' || capability === 'limited'
                }
            };
            await request.patch(requestOptionsWeb);
            if (this.verbose) {
                await logger.logToStderr(`Updating associated member group sharing permissions...`);
            }
            const requestOptionsMemberGroup = {
                url: `${args.options.siteUrl}/_api/Web/AssociatedMemberGroup`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: {
                    AllowMembersEditMembership: capability === 'full'
                }
            };
            await request.patch(requestOptionsMemberGroup);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoSiteSharingPermissionSetCommand();
//# sourceMappingURL=site-sharingpermission-set.js.map