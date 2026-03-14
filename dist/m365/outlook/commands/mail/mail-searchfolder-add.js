import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    userId: z.uuid().optional().alias('i'),
    userName: z.string()
        .refine(userName => validation.isValidUserPrincipalName(userName), {
        error: e => `'${e.input}' is not a valid UPN.`
    }).optional().alias('n'),
    folderName: z.string(),
    messageFilter: z.string(),
    sourceFoldersIds: z.string().transform((value) => value.split(',')).pipe(z.string().array()),
    includeNestedFolders: z.boolean().optional()
});
class OutlookMailSearchFolderAddCommand extends GraphCommand {
    get name() {
        return commands.MAIL_SEARCHFOLDER_ADD;
    }
    get description() {
        return `Creates a new mail search folder in the user's mailbox`;
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => !(options.userId && options.userName), {
            error: 'Specify either userId or userName, but not both'
        });
    }
    async commandAction(logger, args) {
        try {
            const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
            let requestUrl = `${this.resource}/v1.0/me/mailFolders/searchFolders/childFolders`;
            if (isAppOnlyAccessToken) {
                if (!args.options.userId && !args.options.userName) {
                    throw 'When running with application permissions either userId or userName is required';
                }
                const userIdentifier = args.options.userId ?? args.options.userName;
                requestUrl = `${this.resource}/v1.0/users('${userIdentifier}')/mailFolders/searchFolders/childFolders`;
                if (args.options.verbose) {
                    await logger.logToStderr(`Creating a mail search folder in the mailbox of the user ${userIdentifier}...`);
                }
            }
            else {
                if (args.options.userId || args.options.userName) {
                    throw 'You can create mail search folder for other users only if CLI is authenticated in app-only mode';
                }
            }
            const requestOptions = {
                url: requestUrl,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json',
                data: {
                    '@odata.type': '#microsoft.graph.mailSearchFolder',
                    displayName: args.options.folderName,
                    includeNestedFolders: args.options.includeNestedFolders,
                    filterQuery: args.options.messageFilter,
                    sourceFolderIds: args.options.sourceFoldersIds
                }
            };
            const result = await request.post(requestOptions);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new OutlookMailSearchFolderAddCommand();
//# sourceMappingURL=mail-searchfolder-add.js.map