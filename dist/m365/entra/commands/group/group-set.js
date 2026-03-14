var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupSetCommand_instances, _EntraGroupSetCommand_initTelemetry, _EntraGroupSetCommand_initOptions, _EntraGroupSetCommand_initValidators, _EntraGroupSetCommand_initOptionSets, _EntraGroupSetCommand_initTypes;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
class EntraGroupSetCommand extends GraphCommand {
    get name() {
        return commands.GROUP_SET;
    }
    get description() {
        return 'Updates a Microsoft Entra group';
    }
    allowUnknownOptions() {
        return true;
    }
    constructor() {
        super();
        _EntraGroupSetCommand_instances.add(this);
        this.allowedVisibility = ['Public', 'Private'];
        __classPrivateFieldGet(this, _EntraGroupSetCommand_instances, "m", _EntraGroupSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupSetCommand_instances, "m", _EntraGroupSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupSetCommand_instances, "m", _EntraGroupSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraGroupSetCommand_instances, "m", _EntraGroupSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraGroupSetCommand_instances, "m", _EntraGroupSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        let groupId = args.options.id;
        try {
            if (args.options.displayName) {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving group id...`);
                }
                groupId = await entraGroup.getGroupIdByDisplayName(args.options.displayName);
            }
            const requestBody = {
                displayName: args.options.newDisplayName,
                description: args.options.description === '' ? null : args.options.description,
                mailNickName: args.options.mailNickname,
                visibility: args.options.visibility
            };
            this.addUnknownOptionsToPayload(requestBody, args.options);
            const requestOptions = {
                url: `${this.resource}/v1.0/groups/${groupId}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                data: requestBody
            };
            await request.patch(requestOptions);
            const ownerIds = await this.getUserIds(logger, args.options.ownerIds, args.options.ownerUserNames);
            if (ownerIds.length !== 0) {
                await this.updateUsers(logger, groupId, 'owners', ownerIds);
            }
            else if (this.verbose) {
                await logger.logToStderr(`No owners to update.`);
            }
            const memberIds = await this.getUserIds(logger, args.options.memberIds, args.options.memberUserNames);
            if (memberIds.length !== 0) {
                await this.updateUsers(logger, groupId, 'members', memberIds);
            }
            else if (this.verbose) {
                await logger.logToStderr(`No members to update.`);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    ;
    async getUserIds(logger, userIds, userNames) {
        if (userIds) {
            return formatting.splitAndTrim(userIds);
        }
        if (userNames) {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving user IDs...`);
            }
            return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames));
        }
        return [];
    }
    async updateUsers(logger, groupId, role, userIds) {
        const groupUsers = await odata.getAllItems(`${this.resource}/v1.0/groups/${groupId}/${role}/microsoft.graph.user?$select=id`);
        const userIdsToAdd = userIds.filter(userId => !groupUsers.some(groupUser => groupUser.id === userId));
        const userIdsToRemove = groupUsers.filter(groupUser => !userIds.some(userId => groupUser.id === userId)).map(user => user.id);
        if (this.verbose) {
            await logger.logToStderr(`Adding ${userIdsToAdd.length} ${role}...`);
        }
        for (let i = 0; i < userIdsToAdd.length; i += 400) {
            const userIdsBatch = userIdsToAdd.slice(i, i + 400);
            const batchRequestOptions = this.getBatchRequestOptions();
            // only 20 requests per one batch are allowed
            for (let j = 0; j < userIdsBatch.length; j += 20) {
                // only 20 users can be added in one request
                const userIdsChunk = userIdsBatch.slice(j, j + 20);
                batchRequestOptions.data.requests.push({
                    id: j + 1,
                    method: 'PATCH',
                    url: `/groups/${groupId}`,
                    headers: {
                        'content-type': 'application/json;odata.metadata=none',
                        accept: 'application/json;odata.metadata=none'
                    },
                    body: {
                        [`${role}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
                    }
                });
            }
            const res = await request.post(batchRequestOptions);
            for (const response of res.responses) {
                if (response.status !== 204) {
                    throw response.body;
                }
            }
        }
        if (this.verbose) {
            await logger.logToStderr(`Removing ${userIdsToRemove.length} ${role}...`);
        }
        for (let i = 0; i < userIdsToRemove.length; i += 20) {
            const userIdsBatch = userIdsToRemove.slice(i, i + 20);
            const batchRequestOptions = this.getBatchRequestOptions();
            userIdsBatch.map(userId => {
                batchRequestOptions.data.requests.push({
                    id: userId,
                    method: 'DELETE',
                    url: `/groups/${groupId}/${role}/${userId}/$ref`
                });
            });
            const res = await request.post(batchRequestOptions);
            for (const response of res.responses) {
                if (response.status !== 204) {
                    throw response.body;
                }
            }
        }
    }
    getBatchRequestOptions() {
        const requestOptions = {
            url: `${this.resource}/v1.0/$batch`,
            headers: {
                'content-type': 'application/json;odata.metadata=none',
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                requests: []
            }
        };
        return requestOptions;
    }
}
_EntraGroupSetCommand_instances = new WeakSet(), _EntraGroupSetCommand_initTelemetry = function _EntraGroupSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            displayName: typeof args.options.displayName !== 'undefined',
            newDisplayName: typeof args.options.newDisplayName !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            mailNickname: typeof args.options.mailNickname !== 'undefined',
            ownerIds: typeof args.options.ownerIds !== 'undefined',
            ownerUserNames: typeof args.options.ownerUserNames !== 'undefined',
            memberIds: typeof args.options.memberIds !== 'undefined',
            memberUserNames: typeof args.options.memberUserNames !== 'undefined',
            visibility: typeof args.options.visibility !== 'undefined'
        });
        this.trackUnknownOptions(this.telemetryProperties, args.options);
    });
}, _EntraGroupSetCommand_initOptions = function _EntraGroupSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '--mailNickname [mailNickname]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '--newDisplayName [newDisplayName]'
    }, {
        option: '--description [description]'
    }, {
        option: '--ownerIds [ownerIds]'
    }, {
        option: '--ownerUserNames [ownerUserNames]'
    }, {
        option: '--memberIds [memberIds]'
    }, {
        option: '--memberUserNames [memberUserNames]'
    }, {
        option: '--visibility [visibility]',
        autocomplete: this.allowedVisibility
    });
}, _EntraGroupSetCommand_initValidators = function _EntraGroupSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `Value '${args.options.id}' is not a valid GUID for option 'id'.`;
        }
        if (args.options.newDisplayName && args.options.newDisplayName.length > 256) {
            return `The maximum amount of characters for 'newDisplayName' is 256.`;
        }
        if (args.options.mailNickname) {
            if (!validation.isValidMailNickname(args.options.mailNickname)) {
                return `Value '${args.options.mailNickname}' for option 'mailNickname' must contain only characters in the ASCII character set 0-127 except the following: @ () \\ [] " ; : <> , SPACE.`;
            }
            if (args.options.mailNickname.length > 64) {
                return `The maximum amount of characters for 'mailNickname' is 64.`;
            }
        }
        if (args.options.ownerIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.ownerIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for option 'ownerIds': ${isValidGUIDArrayResult}.`;
            }
        }
        if (args.options.ownerUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.ownerUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for option 'ownerUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.memberIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.memberIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for option 'memberIds': ${isValidGUIDArrayResult}.`;
            }
        }
        if (args.options.memberUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.memberUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for option 'memberUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.visibility && !this.allowedVisibility.includes(args.options.visibility)) {
            return `Option 'visibility' must be one of the following values: ${this.allowedVisibility.join(', ')}.`;
        }
        if (args.options.newDisplayName === undefined && args.options.description === undefined && args.options.visibility === undefined
            && args.options.ownerIds === undefined && args.options.ownerUserNames === undefined && args.options.memberIds === undefined
            && args.options.memberUserNames === undefined && args.options.mailNickname === undefined) {
            return `Specify at least one of the following options: 'newDisplayName', 'description', 'visibility', 'ownerIds', 'ownerUserNames', 'memberIds', 'memberUserNames', 'mailNickname'.`;
        }
        return true;
    });
}, _EntraGroupSetCommand_initOptionSets = function _EntraGroupSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName'] }, {
        options: ['ownerIds', 'ownerUserNames'],
        runsWhen: (args) => args.options.ownerIds || args.options.ownerUserNames
    }, {
        options: ['memberIds', 'memberUserNames'],
        runsWhen: (args) => args.options.memberIds || args.options.memberUserNames
    });
}, _EntraGroupSetCommand_initTypes = function _EntraGroupSetCommand_initTypes() {
    this.types.string.push('id', 'displayName', 'newDisplayName', 'description', 'mailNickname', 'ownerIds', 'ownerUserNames', 'memberIds', 'memberUserNames', 'visibility');
};
export default new EntraGroupSetCommand();
//# sourceMappingURL=group-set.js.map