var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupUserSetCommand_instances, _EntraM365GroupUserSetCommand_initTelemetry, _EntraM365GroupUserSetCommand_initOptions, _EntraM365GroupUserSetCommand_initValidators, _EntraM365GroupUserSetCommand_initOptionSets, _EntraM365GroupUserSetCommand_initTypes;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
class EntraM365GroupUserSetCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_USER_SET;
    }
    get description() {
        return 'Updates role of the specified user in the specified Microsoft 365 Group or Microsoft Teams team';
    }
    constructor() {
        super();
        _EntraM365GroupUserSetCommand_instances.add(this);
        this.allowedRoles = ['owner', 'member'];
        __classPrivateFieldGet(this, _EntraM365GroupUserSetCommand_instances, "m", _EntraM365GroupUserSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserSetCommand_instances, "m", _EntraM365GroupUserSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserSetCommand_instances, "m", _EntraM365GroupUserSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserSetCommand_instances, "m", _EntraM365GroupUserSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserSetCommand_instances, "m", _EntraM365GroupUserSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const groupId = await this.getGroupId(logger, args);
            const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);
            if (!isUnifiedGroup) {
                throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
            }
            const userIds = await this.getUserIds(logger, args.options.ids, args.options.userNames);
            // we can't simply switch the role
            // first add users to the new role
            await this.addUsers(groupId, userIds, args.options.role);
            // remove users from the old role
            await this.removeUsersFromRole(logger, groupId, userIds, args.options.role);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getGroupId(logger, args) {
        if (args.options.groupId) {
            return args.options.groupId;
        }
        if (args.options.teamId) {
            return args.options.teamId;
        }
        const name = args.options.groupName || args.options.teamName;
        if (this.verbose) {
            await logger.logToStderr('Retrieving Group ID by display name...');
        }
        return entraGroup.getGroupIdByDisplayName(name);
    }
    async getUserIds(logger, userIds, userNames) {
        if (userIds) {
            return formatting.splitAndTrim(userIds);
        }
        if (this.verbose) {
            await logger.logToStderr('Retrieving user ID(s) by username(s)...');
        }
        return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames));
    }
    async removeUsersFromRole(logger, groupId, userIds, role) {
        const userIdsToRemove = [];
        const currentRole = (role.toLowerCase() === 'member') ? 'owners' : 'members';
        if (this.verbose) {
            await logger.logToStderr(`Removing users from the old role '${currentRole}'.`);
        }
        for (let i = 0; i < userIds.length; i += 20) {
            const userIdsBatch = userIds.slice(i, i + 20);
            const requestOptions = this.getRequestOptions();
            userIdsBatch.map(userId => {
                requestOptions.data.requests.push({
                    id: userId,
                    method: 'GET',
                    url: `/groups/${groupId}/${currentRole}/$count?$filter=id eq '${userId}'`,
                    headers: {
                        'ConsistencyLevel': 'eventual'
                    }
                });
            });
            // send batch request
            const res = await request.post(requestOptions);
            for (const response of res.responses) {
                if (response.status === 200) {
                    if (response.body === 1) {
                        // user can be removed from current role
                        userIdsToRemove.push(response.id);
                    }
                }
                else {
                    throw response.body;
                }
            }
        }
        for (let i = 0; i < userIdsToRemove.length; i += 20) {
            const userIdsBatch = userIds.slice(i, i + 20);
            const requestOptions = this.getRequestOptions();
            userIdsBatch.map(userId => {
                requestOptions.data.requests.push({
                    id: userId,
                    method: 'DELETE',
                    url: `/groups/${groupId}/${currentRole}/${userId}/$ref`
                });
            });
            const res = await request.post(requestOptions);
            for (const response of res.responses) {
                if (response.status !== 204) {
                    throw response.body;
                }
            }
        }
    }
    async addUsers(groupId, userIds, role) {
        for (let i = 0; i < userIds.length; i += 400) {
            const userIdsBatch = userIds.slice(i, i + 400);
            const requestOptions = this.getRequestOptions();
            for (let j = 0; j < userIdsBatch.length; j += 20) {
                const userIdsChunk = userIdsBatch.slice(j, j + 20);
                requestOptions.data.requests.push({
                    id: j + 1,
                    method: 'PATCH',
                    url: `/groups/${groupId}`,
                    headers: {
                        'content-type': 'application/json;odata.metadata=none'
                    },
                    body: {
                        [`${role.toLowerCase() === 'member' ? 'members' : 'owners'}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
                    }
                });
            }
            const res = await request.post(requestOptions);
            for (const response of res.responses) {
                if (response.status !== 204) {
                    throw response.body;
                }
            }
        }
    }
    getRequestOptions() {
        const requestOptions = {
            url: `${this.resource}/v1.0/$batch`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                requests: []
            }
        };
        return requestOptions;
    }
}
_EntraM365GroupUserSetCommand_instances = new WeakSet(), _EntraM365GroupUserSetCommand_initTelemetry = function _EntraM365GroupUserSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            ids: typeof args.options.ids !== 'undefined',
            userNames: typeof args.options.userNames !== 'undefined'
        });
    });
}, _EntraM365GroupUserSetCommand_initOptions = function _EntraM365GroupUserSetCommand_initOptions() {
    this.options.unshift({
        option: '--ids [ids]'
    }, {
        option: '--userNames [userNames]'
    }, {
        option: '-i, --groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '--teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '-r, --role <role>',
        autocomplete: this.allowedRoles
    });
}, _EntraM365GroupUserSetCommand_initValidators = function _EntraM365GroupUserSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `'${args.options.teamId}' is not a valid GUID for option 'teamId'.`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `'${args.options.groupId}' is not a valid GUID for option 'groupId'.`;
        }
        if (args.options.ids) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.ids);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'ids': ${isValidGUIDArrayResult}.`;
            }
        }
        if (args.options.userNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.userNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'userNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.role && !this.allowedRoles.some(role => role.toLowerCase() === args.options.role.toLowerCase())) {
            return `'${args.options.role}' is not a valid role. Allowed values are: ${this.allowedRoles.join(',')}`;
        }
        return true;
    });
}, _EntraM365GroupUserSetCommand_initOptionSets = function _EntraM365GroupUserSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['groupId', 'groupName', 'teamId', 'teamName'] });
    this.optionSets.push({ options: ['ids', 'userNames'] });
}, _EntraM365GroupUserSetCommand_initTypes = function _EntraM365GroupUserSetCommand_initTypes() {
    this.types.string.push('ids', 'userNames', 'groupId', 'groupName', 'teamId', 'teamName', 'role');
};
export default new EntraM365GroupUserSetCommand();
//# sourceMappingURL=m365group-user-set.js.map