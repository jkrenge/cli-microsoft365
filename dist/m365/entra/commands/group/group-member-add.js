var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupMemberAddCommand_instances, _EntraGroupMemberAddCommand_initTelemetry, _EntraGroupMemberAddCommand_initOptions, _EntraGroupMemberAddCommand_initValidators, _EntraGroupMemberAddCommand_initOptionSets, _EntraGroupMemberAddCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupMemberAddCommand extends GraphCommand {
    get name() {
        return commands.GROUP_MEMBER_ADD;
    }
    get description() {
        return 'Adds members to a Microsoft Entra group';
    }
    constructor() {
        super();
        _EntraGroupMemberAddCommand_instances.add(this);
        this.roleValues = ['Owner', 'Member'];
        __classPrivateFieldGet(this, _EntraGroupMemberAddCommand_instances, "m", _EntraGroupMemberAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberAddCommand_instances, "m", _EntraGroupMemberAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberAddCommand_instances, "m", _EntraGroupMemberAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberAddCommand_instances, "m", _EntraGroupMemberAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberAddCommand_instances, "m", _EntraGroupMemberAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Adding member(s) ${args.options.userIds || args.options.userNames || args.options.subgroupIds || args.options.subgroupNames} to group ${args.options.groupId || args.options.groupName}...`);
            }
            const groupId = await this.getGroupId(logger, args.options);
            const objectIds = await this.getObjectIds(logger, args.options);
            for (let i = 0; i < objectIds.length; i += 400) {
                const objectIdsBatch = objectIds.slice(i, i + 400);
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
                for (let j = 0; j < objectIdsBatch.length; j += 20) {
                    const objectIdsChunk = objectIdsBatch.slice(j, j + 20);
                    requestOptions.data.requests.push({
                        id: j + 1,
                        method: 'PATCH',
                        url: `/groups/${groupId}`,
                        headers: {
                            'content-type': 'application/json;odata.metadata=none'
                        },
                        body: {
                            [`${args.options.role === 'Member' ? 'members' : 'owners'}@odata.bind`]: objectIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
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
    async getObjectIds(logger, options) {
        if (options.userIds || options.userNames) {
            return this.getUserIds(logger, options);
        }
        return this.getGroupIds(logger, options);
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
    async getGroupIds(logger, options) {
        if (options.subgroupIds) {
            return options.subgroupIds.split(',').map(i => i.trim());
        }
        if (this.verbose) {
            await logger.logToStderr('Retrieving ID(s) of group(s)...');
        }
        return entraGroup.getGroupIdsByDisplayNames(options.subgroupNames.split(',').map(u => u.trim()));
    }
}
_EntraGroupMemberAddCommand_instances = new WeakSet(), _EntraGroupMemberAddCommand_initTelemetry = function _EntraGroupMemberAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            userIds: typeof args.options.userIds !== 'undefined',
            userNames: typeof args.options.userNames !== 'undefined',
            subgroupIds: typeof args.options.subgroupIds !== 'undefined',
            subgroupNames: typeof args.options.subgroupNames !== 'undefined'
        });
    });
}, _EntraGroupMemberAddCommand_initOptions = function _EntraGroupMemberAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --groupId [groupId]'
    }, {
        option: '-n, --groupName [groupName]'
    }, {
        option: '--userIds [userIds]'
    }, {
        option: '--userNames [userNames]'
    }, {
        option: '--subgroupIds [subgroupIds]'
    }, {
        option: '--subgroupNames [subgroupNames]'
    }, {
        option: '-r, --role <role>',
        autocomplete: this.roleValues
    });
}, _EntraGroupMemberAddCommand_initValidators = function _EntraGroupMemberAddCommand_initValidators() {
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
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.userNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'userNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.subgroupIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.subgroupIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'subgroupIds': ${isValidGUIDArrayResult}.`;
            }
        }
        if ((args.options.subgroupIds || args.options.subgroupNames) && args.options.role === 'Owner') {
            return `Subgroups cannot be set as owners.`;
        }
        if (this.roleValues.indexOf(args.options.role) === -1) {
            return `Option 'role' must be one of the following values: ${this.roleValues.join(', ')}.`;
        }
        return true;
    });
}, _EntraGroupMemberAddCommand_initOptionSets = function _EntraGroupMemberAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['groupId', 'groupName'] }, { options: ['userIds', 'userNames', 'subgroupIds', 'subgroupNames'] });
}, _EntraGroupMemberAddCommand_initTypes = function _EntraGroupMemberAddCommand_initTypes() {
    this.types.string.push('groupId', 'groupName', 'ids', 'userIds', 'userNames', 'subgroupIds', 'subgroupNames', 'role');
};
export default new EntraGroupMemberAddCommand();
//# sourceMappingURL=group-member-add.js.map