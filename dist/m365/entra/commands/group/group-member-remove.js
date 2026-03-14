var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupMemberRemoveCommand_instances, _EntraGroupMemberRemoveCommand_initTelemetry, _EntraGroupMemberRemoveCommand_initOptions, _EntraGroupMemberRemoveCommand_initValidators, _EntraGroupMemberRemoveCommand_initOptionSets, _EntraGroupMemberRemoveCommand_initTypes;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
class EntraGroupMemberRemoveCommand extends GraphCommand {
    get name() {
        return commands.GROUP_MEMBER_REMOVE;
    }
    get description() {
        return 'Removes members from a Microsoft Entra group';
    }
    constructor() {
        super();
        _EntraGroupMemberRemoveCommand_instances.add(this);
        this.roleValues = ['Owner', 'Member'];
        __classPrivateFieldGet(this, _EntraGroupMemberRemoveCommand_instances, "m", _EntraGroupMemberRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberRemoveCommand_instances, "m", _EntraGroupMemberRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberRemoveCommand_instances, "m", _EntraGroupMemberRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberRemoveCommand_instances, "m", _EntraGroupMemberRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraGroupMemberRemoveCommand_instances, "m", _EntraGroupMemberRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const removeUsers = async () => {
                if (this.verbose) {
                    await logger.logToStderr(`Removing user(s) ${args.options.userIds || args.options.userNames || args.options.subgroupIds || args.options.subgroupNames} from group ${args.options.groupId || args.options.groupName}...`);
                }
                const groupId = await this.getGroupId(logger, args.options);
                const userIds = await this.getPrincipalIds(logger, args.options);
                const endpoints = [];
                if (!args.options.role || args.options.role === 'Owner') {
                    endpoints.push(...userIds.map(id => `/groups/${groupId}/owners/${id}/$ref`));
                }
                if (!args.options.role || args.options.role === 'Member') {
                    endpoints.push(...userIds.map(id => `/groups/${groupId}/members/${id}/$ref`));
                }
                for (let i = 0; i < endpoints.length; i += 20) {
                    const endpointsBatch = endpoints.slice(i, i + 20);
                    const requestOptions = {
                        url: `${this.resource}/v1.0/$batch`,
                        headers: {
                            'content-type': 'application/json;odata.metadata=none'
                        },
                        responseType: 'json',
                        data: {
                            requests: endpointsBatch.map((ep, index) => ({
                                id: index + 1,
                                method: 'DELETE',
                                url: ep,
                                headers: {
                                    'content-type': 'application/json;odata.metadata=none'
                                }
                            }))
                        }
                    };
                    const res = await request.post(requestOptions);
                    for (const response of res.responses) {
                        // Suppress 404 errors if suppressNotFound is set
                        if (response.status !== 204 && (!args.options.suppressNotFound || response.status !== 404)) {
                            throw response.body;
                        }
                    }
                }
            };
            if (args.options.force) {
                await removeUsers();
            }
            else {
                const principals = args.options.userIds || args.options.userNames || args.options.subgroupIds || args.options.subgroupNames;
                const principalsList = principals.split(',');
                const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove ${principalsList.length} principal(s) from group '${args.options.groupId || args.options.groupName}'?` });
                if (result) {
                    await removeUsers();
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
            await logger.logToStderr(`Retrieving ID of group '${options.groupName}'...`);
        }
        return entraGroup.getGroupIdByDisplayName(options.groupName);
    }
    async getPrincipalIds(logger, options) {
        if (options.userIds) {
            return options.userIds.split(',').map(i => i.trim());
        }
        if (options.subgroupIds) {
            return options.subgroupIds.split(',').map(i => i.trim());
        }
        if (options.userNames) {
            if (this.verbose) {
                await logger.logToStderr('Retrieving ID(s) of user(s)...');
            }
            return entraUser.getUserIdsByUpns(options.userNames.split(',').map(u => u.trim()));
        }
        // Subgroup names were specified
        if (this.verbose) {
            await logger.logToStderr('Retrieving ID(s) of subgroup(s)...');
        }
        const subGroupIds = [];
        for (const subgroupName of options.subgroupNames.split(',')) {
            const groupId = await entraGroup.getGroupIdByDisplayName(subgroupName.trim());
            subGroupIds.push(groupId);
        }
        return subGroupIds;
    }
}
_EntraGroupMemberRemoveCommand_instances = new WeakSet(), _EntraGroupMemberRemoveCommand_initTelemetry = function _EntraGroupMemberRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            userIds: typeof args.options.userIds !== 'undefined',
            userNames: typeof args.options.userNames !== 'undefined',
            subgroupIds: typeof args.options.subgroupIds !== 'undefined',
            subgroupNames: typeof args.options.subgroupNames !== 'undefined',
            role: typeof args.options.role !== 'undefined',
            suppressNotFound: !!args.options.suppressNotFound,
            force: !!args.options.force
        });
    });
}, _EntraGroupMemberRemoveCommand_initOptions = function _EntraGroupMemberRemoveCommand_initOptions() {
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
        option: '-r, --role [role]',
        autocomplete: this.roleValues
    }, {
        option: '--suppressNotFound'
    }, {
        option: '-f, --force'
    });
}, _EntraGroupMemberRemoveCommand_initValidators = function _EntraGroupMemberRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.groupId !== undefined && !validation.isValidGuid(args.options.groupId)) {
            return `'${args.options.groupId}' is not a valid GUID for option 'groupId'.`;
        }
        if (args.options.userIds !== undefined) {
            const invalidGuids = validation.isValidGuidArray(args.options.userIds);
            if (invalidGuids !== true) {
                return `Invalid GUIDs found for option 'ids': ${invalidGuids}.`;
            }
        }
        if (args.options.userNames !== undefined) {
            const invalidUpns = validation.isValidUserPrincipalNameArray(args.options.userNames);
            if (invalidUpns !== true) {
                return `Invalid UPNs found for option 'userNames': ${invalidUpns}.`;
            }
        }
        if (args.options.subgroupIds !== undefined) {
            const invalidGuids = validation.isValidGuidArray(args.options.subgroupIds);
            if (invalidGuids !== true) {
                return `Invalid GUIDs found for option 'subgroupIds': ${invalidGuids}.`;
            }
        }
        if (args.options.role !== undefined && this.roleValues.indexOf(args.options.role) === -1) {
            return `Option 'role' must be one of the following values: ${this.roleValues.join(', ')}.`;
        }
        if ((args.options.subgroupIds !== undefined || args.options.subgroupNames !== undefined) && args.options.role?.toLowerCase() !== 'member') {
            return `When removing subgroups, the 'role' option must be set to 'Member'.`;
        }
        return true;
    });
}, _EntraGroupMemberRemoveCommand_initOptionSets = function _EntraGroupMemberRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['groupId', 'groupName'] }, { options: ['userIds', 'userNames', 'subgroupIds', 'subgroupNames'] });
}, _EntraGroupMemberRemoveCommand_initTypes = function _EntraGroupMemberRemoveCommand_initTypes() {
    this.types.string.push('groupId', 'groupName', 'ids', 'userNames', 'subgroupIds', 'subgroupNames', 'role');
    this.types.boolean.push('force', 'suppressNotFound');
};
export default new EntraGroupMemberRemoveCommand();
//# sourceMappingURL=group-member-remove.js.map