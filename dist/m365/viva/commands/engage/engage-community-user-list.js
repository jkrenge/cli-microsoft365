import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { odata } from '../../../../utils/odata.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    communityId: z.string().optional(),
    communityDisplayName: z.string().optional().alias('n'),
    entraGroupId: z.uuid().optional(),
    role: z.enum(['Admin', 'Member']).optional().alias('r')
});
class VivaEngageCommunityUserListCommand extends GraphCommand {
    get name() {
        return commands.ENGAGE_COMMUNITY_USER_LIST;
    }
    get description() {
        return 'Lists all users within a specified Microsoft 365 Viva Engage community';
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
        });
    }
    defaultProperties() {
        return ['id', 'displayName', 'userPrincipalName', 'roles'];
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr('Getting list of users in community...');
            }
            let entraGroupId = args.options.entraGroupId;
            if (args.options.communityDisplayName) {
                const community = await vivaEngage.getCommunityByDisplayName(args.options.communityDisplayName, ['groupId']);
                entraGroupId = community.groupId;
            }
            if (args.options.communityId) {
                const community = await vivaEngage.getCommunityById(args.options.communityId, ['groupId']);
                entraGroupId = community.groupId;
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/groups/${entraGroupId}/members`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const members = await odata.getAllItems(requestOptions);
            requestOptions.url = `${this.resource}/v1.0/groups/${entraGroupId}/owners`;
            const owners = await odata.getAllItems(requestOptions);
            const extendedMembers = members.map(m => {
                return {
                    ...m,
                    roles: ['Member']
                };
            });
            const extendedOwners = owners.map(o => {
                return {
                    ...o,
                    roles: ['Admin']
                };
            });
            let users = [];
            if (args.options.role) {
                if (args.options.role === 'Member') {
                    users = users.concat(extendedMembers);
                }
                if (args.options.role === 'Admin') {
                    users = users.concat(extendedOwners);
                }
            }
            else {
                users = extendedOwners.concat(extendedMembers);
            }
            await logger.log(users);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new VivaEngageCommunityUserListCommand();
//# sourceMappingURL=engage-community-user-list.js.map