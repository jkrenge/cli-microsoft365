import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    userId: z.uuid().optional().alias('i'),
    userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
        error: e => `'${e.input}' is not a valid UPN.`
    }).optional().alias('n'),
    force: z.boolean().optional().alias('f')
});
class EntraUserSessionRevokeCommand extends GraphCommand {
    get name() {
        return commands.USER_SESSION_REVOKE;
    }
    get description() {
        return 'Revokes all sign-in sessions for a given user';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.userId, options.userName].filter(o => o !== undefined).length === 1, {
            error: `Specify either 'userId' or 'userName'.`
        });
    }
    async commandAction(logger, args) {
        const revokeUserSessions = async () => {
            try {
                const userIdentifier = args.options.userId ?? args.options.userName;
                if (this.verbose) {
                    await logger.logToStderr(`Invalidating all the refresh tokens for user ${userIdentifier}...`);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/users('${formatting.encodeQueryParameter(userIdentifier)}')/revokeSignInSessions`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json',
                    data: {}
                };
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await revokeUserSessions();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `This will revoke all sessions for the user '${args.options.userId || args.options.userName}', requiring the user to re-sign in from all devices. Are you sure?` });
            if (result) {
                await revokeUserSessions();
            }
        }
    }
}
export default new EntraUserSessionRevokeCommand();
//# sourceMappingURL=user-session-revoke.js.map