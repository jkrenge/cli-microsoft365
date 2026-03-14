import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    url: z.string()
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    }).alias('u'),
    audienceIds: z.string()
        .refine(audiences => validation.isValidGuidArray(audiences) === true, {
        error: e => `The following GUIDs are invalid: ${e.input}.`
    }).optional(),
    audienceNames: z.string().optional(),
    vivaConnectionsDefaultStart: z.boolean().optional(),
    isInDraftMode: z.boolean().optional(),
    order: z.number()
        .refine(order => validation.isValidPositiveInteger(order) === true, {
        error: e => `'${e.input}' is not a positive integer.`
    }).optional()
});
class SpoHomeSiteAddCommand extends SpoCommand {
    get name() {
        return commands.HOMESITE_ADD;
    }
    get description() {
        return 'Adds a home site';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine((options) => [options.audienceIds, options.audienceNames].filter(o => o !== undefined).length <= 1, {
            message: 'You must specify either audienceIds or audienceNames but not both.'
        });
    }
    async commandAction(logger, args) {
        let audiences = [];
        if (args.options.audienceIds) {
            audiences = args.options.audienceIds.split(',').map(id => id.trim());
        }
        else if (args.options.audienceNames) {
            audiences = await this.transformAudienceNamesToIds(args.options.audienceNames);
        }
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);
            const requestOptions = {
                url: `${spoAdminUrl}/_api/SPHSite/AddHomeSite`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: {
                    siteUrl: args.options.url,
                    audiences: audiences,
                    vivaConnectionsDefaultStart: args.options.vivaConnectionsDefaultStart ?? true,
                    isInDraftMode: args.options.isInDraftMode ?? false,
                    order: args.options.order
                }
            };
            if (this.verbose) {
                await logger.logToStderr(`Adding home site with URL: ${args.options.url}...`);
            }
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async transformAudienceNamesToIds(audienceNames) {
        const names = audienceNames.split(',');
        const ids = [];
        for (const name of names) {
            const id = await entraGroup.getGroupIdByDisplayName(name.trim());
            ids.push(id);
        }
        return ids;
    }
}
export default new SpoHomeSiteAddCommand();
//# sourceMappingURL=homesite-add.js.map