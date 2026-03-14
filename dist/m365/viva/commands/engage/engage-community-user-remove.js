import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { cli } from '../../../../cli/cli.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    communityId: z.string().optional(),
    communityDisplayName: z.string().optional().alias('n'),
    entraGroupId: z.uuid().optional(),
    id: z.uuid().optional(),
    userName: z.string()
        .refine(userName => validation.isValidUserPrincipalName(userName), {
        error: e => `'${e.input}' is not a valid user principal name.`
    }).optional(),
    force: z.boolean().optional()
});
class VivaEngageCommunityUserRemoveCommand extends GraphCommand {
    get name() {
        return commands.ENGAGE_COMMUNITY_USER_REMOVE;
    }
    get description() {
        return 'Removes a specified user from a Microsoft 365 Viva Engage community';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => [options.communityId, options.communityDisplayName, options.entraGroupId].filter(x => x !== undefined).length === 1, {
            error: 'Specify either communityId, communityDisplayName, or entraGroupId, but not multiple.'
        })
            .refine(options => options.communityId || options.communityDisplayName || options.entraGroupId, {
            error: 'Specify at least one of communityId, communityDisplayName, or entraGroupId.'
        })
            .refine(options => options.id || options.userName, {
            error: 'Specify either of id or userName.'
        })
            .refine(options => typeof options.userName !== 'undefined' || typeof options.id !== 'undefined', {
            error: 'Specify either id or userName, but not both.'
        });
    }
    async commandAction(logger, args) {
        try {
            if (args.options.force) {
                await this.deleteUserFromCommunity(args.options, logger);
            }
            else {
                const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the user ${args.options.id || args.options.userName} from the community ${args.options.communityDisplayName || args.options.communityId || args.options.entraGroupId}?` });
                if (result) {
                    await this.deleteUserFromCommunity(args.options, logger);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async deleteUserFromCommunity(options, logger) {
        if (this.verbose) {
            await logger.logToStderr('Removing user from community...');
        }
        let entraGroupId = options.entraGroupId;
        if (options.communityDisplayName) {
            const community = await vivaEngage.getCommunityByDisplayName(options.communityDisplayName, ['groupId']);
            entraGroupId = community.groupId;
        }
        else if (options.communityId) {
            const community = await vivaEngage.getCommunityById(options.communityId, ['groupId']);
            entraGroupId = community.groupId;
        }
        const userId = options.id || await entraUser.getUserIdByUpn(options.userName);
        await this.deleteUser(entraGroupId, userId, 'owners');
        await this.deleteUser(entraGroupId, userId, 'members');
    }
    async deleteUser(entraGroupId, userId, role) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/groups/${entraGroupId}/${role}/${userId}/$ref`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            if (err.response.status !== 404) {
                throw err.response.data;
            }
        }
    }
}
export default new VivaEngageCommunityUserRemoveCommand();
//# sourceMappingURL=engage-community-user-remove.js.map