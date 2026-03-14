import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    userId: z.uuid().optional().alias('i'),
    userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
        error: e => `'${e.input}' is not a valid UPN.`
    }).optional().alias('n')
});
class OutlookMailboxSettingsGetCommand extends GraphCommand {
    get name() {
        return commands.MAILBOX_SETTINGS_GET;
    }
    get description() {
        return `Get the user's mailbox settings`;
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
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
        let requestUrl = `${this.resource}/v1.0/me/mailboxSettings`;
        if (isAppOnlyAccessToken) {
            if (!args.options.userId && !args.options.userName) {
                throw 'When running with application permissions either userId or userName is required';
            }
            const userIdentifier = args.options.userId ?? args.options.userName;
            if (this.verbose) {
                await logger.logToStderr(`Retrieving mailbox settings for user ${userIdentifier}...`);
            }
            requestUrl = `${this.resource}/v1.0/users('${userIdentifier}')/mailboxSettings`;
        }
        else {
            if (args.options.userId || args.options.userName) {
                throw 'You can retrieve mailbox settings of other users only if CLI is authenticated in app-only mode';
            }
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const result = await request.get(requestOptions);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new OutlookMailboxSettingsGetCommand();
//# sourceMappingURL=mailbox-settings-get.js.map