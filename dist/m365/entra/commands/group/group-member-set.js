var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupMemberSetCommand_instances, _EntraGroupMemberSetCommand_initTelemetry, _EntraGroupMemberSetCommand_initOptions, _EntraGroupMemberSetCommand_initValidators, _EntraGroupMemberSetCommand_initOptionSets, _EntraGroupMemberSetCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupMemberSetCommand extends GraphCommand {
    get name() {
        return commands.GROUP_MEMBER_SET;
    }
    get description() {
        return 'Updates the role of members in a Microsoft Entra ID group';
    }
    constructor() {
        super();
        _EntraGroupMemberSetCommand_instances.add(this);
        this.roleValues = ['Owner', 'Member'];
        __classPrivateFieldGet(this, _EntraGroupMemberSetCommand_instances, "m", _EntraGroupMemberSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberSetCommand_instances, "m", _EntraGroupMemberSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberSetCommand_instances, "m", _EntraGroupMemberSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberSetCommand_instances, "m", _EntraGroupMemberSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberSetCommand_instances, "m", _EntraGroupMemberSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Adding member(s) ${args.options.userIds || args.options.userNames} to role ${args.options.role} of group ${args.options.groupId || args.options.groupName}...`);
            }
            const groupId = await this.getGroupId(logger, args.options);
            const userIds = await this.getUserIds(logger, args.options);
            // we can't simply switch the role
            // first add users to the new role
            await this.addUsers(groupId, userIds, args.options);
            // remove users from the old role
            await this.removeUsersFromRole(logger, groupId, userIds, args.options);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getGroupId(logger, options) {
        if (options.groupId) {
            return options.groupId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving ID of group ${options.groupName}...`);
        }
        return entraGroup.getGroupIdByDisplayName(options.groupName);
    }
    async getUserIds(logger, options) {
        if (options.userIds) {
            return options.userIds.split(',').map(i => i.trim());
        }
        if (this.verbose) {
            await logger.logToStderr('Retrieving ID(s) of user(s)...');
        }
        return entraUser.getUserIdsByUpns(options.userNames.split(',').map(u => u.trim()));
    }
    async removeUsersFromRole(logger, groupId, userIds, options) {
        const userIdsToRemove = [];
        const currentRole = options.role === 'Member' ? 'owners' : 'members';
        if (this.verbose) {
            await logger.logToStderr(`Removing members from the old role '${currentRole}'.`);
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
    async addUsers(groupId, userIds, options) {
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
                        [`${options.role === 'Member' ? 'members' : 'owners'}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
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
_EntraGroupMemberSetCommand_instances = new WeakSet(), _EntraGroupMemberSetCommand_initTelemetry = function _EntraGroupMemberSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            userIds: typeof args.options.userIds !== 'undefined',
            userNames: typeof args.options.userNames !== 'undefined'
        });
    });
}, _EntraGroupMemberSetCommand_initOptions = function _EntraGroupMemberSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --groupId [groupId]'
    }, {
        option: '-n, --groupName [groupName]'
    }, {
        option: '--userIds [userIds]'
    }, {
        option: '--userNames [userNames]'
    }, {
        option: '-r, --role <role>',
        autocomplete: this.roleValues
    });
}, _EntraGroupMemberSetCommand_initValidators = function _EntraGroupMemberSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID for option groupId.`;
        }
        if (args.options.userIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.userIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'userIds': ${isValidGUIDArrayResult}.`;
            }
        }
        if (args.options.userNames) {
            const isValidUserPrincipalNameArray = validation.isValidUserPrincipalNameArray(args.options.userNames);
            if (isValidUserPrincipalNameArray !== true) {
                return `User principal name '${isValidUserPrincipalNameArray}' is invalid for option 'userNames'.`;
            }
        }
        if (this.roleValues.indexOf(args.options.role) === -1) {
            return `Option 'role' must be one of the following values: ${this.roleValues.join(', ')}.`;
        }
        return true;
    });
}, _EntraGroupMemberSetCommand_initOptionSets = function _EntraGroupMemberSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['groupId', 'groupName'] }, { options: ['userIds', 'userNames'] });
}, _EntraGroupMemberSetCommand_initTypes = function _EntraGroupMemberSetCommand_initTypes() {
    this.types.string.push('groupId', 'groupName', 'userIds', 'userNames', 'role');
};
export default new EntraGroupMemberSetCommand();
//# sourceMappingURL=group-member-set.js.map