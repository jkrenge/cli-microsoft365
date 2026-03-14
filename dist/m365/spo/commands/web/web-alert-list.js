import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { globalOptionsZod } from '../../../../Command.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    webUrl: z.string().alias('u')
        .refine(url => validation.isValidSharePointUrl(url) === true, {
        error: e => `'${e.input}' is not a valid SharePoint site URL.`
    }),
    listId: z.uuid().optional(),
    listUrl: z.string().optional(),
    listTitle: z.string().optional(),
    userName: z.string().optional().refine(upn => typeof upn === 'undefined' || validation.isValidUserPrincipalName(upn), {
        error: e => `'${e.input}' is not a valid UPN.`
    }),
    userId: z.uuid().optional()
});
class SpoWebAlertListCommand extends SpoCommand {
    get name() {
        return commands.WEB_ALERT_LIST;
    }
    get description() {
        return 'Lists all SharePoint list alerts';
    }
    defaultProperties() {
        return ['ID', 'Title', 'UserPrincipalName'];
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.listId, options.listUrl, options.listTitle].filter(x => x !== undefined).length <= 1, {
            error: `Specify either listId, listUrl, or listTitle, but not more than one.`
        })
            .refine(options => [options.userName, options.userId].filter(x => x !== undefined).length <= 1, {
            error: `Specify either userName or userId, but not both.`
        });
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            const listParams = args.options.listId || args.options.listTitle || args.options.listUrl;
            const userParams = args.options.userName || args.options.userId;
            let message = `Retrieving alerts from site '${args.options.webUrl}'`;
            if (listParams) {
                message += ` for list '${listParams}'`;
            }
            if (userParams) {
                message += `${listParams ? ' and' : ''} for user '${userParams}'`;
            }
            await logger.logToStderr(`${message}...`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web/alerts?$expand=List,User,List/Rootfolder,Item&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl,Item/ID,Item/FileRef,Item/Guid`;
        const filters = [];
        let listId;
        if (args.options.listId) {
            listId = args.options.listId;
        }
        else if (args.options.listUrl || args.options.listTitle) {
            listId = await spo.getListId(args.options.webUrl, args.options.listTitle, args.options.listUrl, logger, this.verbose);
        }
        if (listId) {
            filters.push(`List/Id eq guid'${formatting.encodeQueryParameter(listId)}'`);
        }
        if (args.options.userName) {
            filters.push(`User/UserPrincipalName eq '${formatting.encodeQueryParameter(args.options.userName)}'`);
        }
        else if (args.options.userId) {
            const userPrincipalName = await entraUser.getUpnByUserId(args.options.userId);
            filters.push(`User/UserPrincipalName eq '${formatting.encodeQueryParameter(userPrincipalName)}'`);
        }
        if (filters.length > 0) {
            requestUrl += `&$filter=${filters.join(' and ')}`;
        }
        try {
            const res = await odata.getAllItems(requestUrl);
            res.forEach(alert => {
                if (alert.Item) {
                    delete alert.Item['ID'];
                }
                if (cli.shouldTrimOutput(args.options.output)) {
                    alert.UserPrincipalName = alert.User?.UserPrincipalName;
                }
            });
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoWebAlertListCommand();
//# sourceMappingURL=web-alert-list.js.map