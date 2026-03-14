var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraM365GroupUserRemoveCommand_instances, _EntraM365GroupUserRemoveCommand_initTelemetry, _EntraM365GroupUserRemoveCommand_initOptions, _EntraM365GroupUserRemoveCommand_initValidators, _EntraM365GroupUserRemoveCommand_initOptionSets, _EntraM365GroupUserRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraM365GroupUserRemoveCommand extends GraphCommand {
    get name() {
        return commands.M365GROUP_USER_REMOVE;
    }
    get description() {
        return 'Removes the specified user from specified Microsoft 365 Group or Microsoft Teams team';
    }
    constructor() {
        super();
        _EntraM365GroupUserRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserRemoveCommand_instances, "m", _EntraM365GroupUserRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserRemoveCommand_instances, "m", _EntraM365GroupUserRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserRemoveCommand_instances, "m", _EntraM365GroupUserRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserRemoveCommand_instances, "m", _EntraM365GroupUserRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraM365GroupUserRemoveCommand_instances, "m", _EntraM365GroupUserRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeUser = async () => {
            try {
                const groupId = await this.getGroupId(logger, args.options.groupId, args.options.teamId, args.options.groupName, args.options.teamName);
                const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);
                if (!isUnifiedGroup) {
                    throw Error(`Specified group with id '${groupId}' is not a Microsoft 365 group.`);
                }
                const userNames = args.options.userNames;
                const userIds = await this.getUserIds(logger, args.options.ids, userNames);
                await this.removeUsersFromGroup(groupId, userIds, 'owners');
                await this.removeUsersFromGroup(groupId, userIds, 'members');
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeUser();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove ${args.options.userNames || args.options.ids} from ${args.options.groupId || args.options.groupName || args.options.teamId || args.options.teamName}?` });
            if (result) {
                await removeUser();
            }
        }
    }
    async getGroupId(logger, groupId, teamId, groupName, teamName) {
        const id = groupId || teamId;
        if (id) {
            return id;
        }
        const name = groupName ?? teamName;
        if (this.verbose) {
            await logger.logToStderr(`Retrieving Group ID by display name ${name}...`);
        }
        return entraGroup.getGroupIdByDisplayName(name);
    }
    async getUserIds(logger, userIds, userNames) {
        if (userIds) {
            return formatting.splitAndTrim(userIds);
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving user IDs for {userNames}...`);
        }
        return entraUser.getUserIdsByUpns(formatting.splitAndTrim(userNames));
    }
    async removeUsersFromGroup(groupId, userIds, role) {
        for (const userId of userIds) {
            try {
                await request.delete({
                    url: `${this.resource}/v1.0/groups/${groupId}/${role}/${userId}/$ref`,
                    headers: {
                        'accept': 'application/json;odata.metadata=none'
                    }
                });
            }
            catch (err) {
                // the 404 error is accepted
                if (err.response.status !== 404) {
                    throw err.response.data;
                }
            }
        }
    }
}
_EntraM365GroupUserRemoveCommand_instances = new WeakSet(), _EntraM365GroupUserRemoveCommand_initTelemetry = function _EntraM365GroupUserRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force,
            teamId: typeof args.options.teamId !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            ids: typeof args.options.ids !== 'undefined',
            userNames: typeof args.options.userNames !== 'undefined'
        });
    });
}, _EntraM365GroupUserRemoveCommand_initOptions = function _EntraM365GroupUserRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '--teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '--ids [ids]'
    }, {
        option: '--userNames [userNames]'
    }, {
        option: '-f, --force'
    });
}, _EntraM365GroupUserRemoveCommand_initValidators = function _EntraM365GroupUserRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID for option 'teamId'.`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID for option 'groupId'.`;
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
        return true;
    });
}, _EntraM365GroupUserRemoveCommand_initOptionSets = function _EntraM365GroupUserRemoveCommand_initOptionSets() {
    this.optionSets.push({
        options: ['groupId', 'teamId', 'groupName', 'teamName']
    }, {
        options: ['ids', 'userNames']
    });
}, _EntraM365GroupUserRemoveCommand_initTypes = function _EntraM365GroupUserRemoveCommand_initTypes() {
    this.types.string.push('groupId', 'groupName', 'teamId', 'teamName', 'ids', 'userNames');
    this.types.boolean.push('force');
};
export default new EntraM365GroupUserRemoveCommand();
//# sourceMappingURL=m365group-user-remove.js.map