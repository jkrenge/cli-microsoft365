var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraPimRoleAssignmentRemoveCommand_instances, _EntraPimRoleAssignmentRemoveCommand_initTelemetry, _EntraPimRoleAssignmentRemoveCommand_initOptions, _EntraPimRoleAssignmentRemoveCommand_initValidators, _EntraPimRoleAssignmentRemoveCommand_initOptionSets, _EntraPimRoleAssignmentRemoveCommand_initTypes;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { validation } from '../../../../utils/validation.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
class EntraPimRoleAssignmentRemoveCommand extends GraphCommand {
    get name() {
        return commands.PIM_ROLE_ASSIGNMENT_REMOVE;
    }
    get description() {
        return 'Request deactivation of an Entra role assignment for a user or group';
    }
    constructor() {
        super();
        _EntraPimRoleAssignmentRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentRemoveCommand_instances, "m", _EntraPimRoleAssignmentRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentRemoveCommand_instances, "m", _EntraPimRoleAssignmentRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentRemoveCommand_instances, "m", _EntraPimRoleAssignmentRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentRemoveCommand_instances, "m", _EntraPimRoleAssignmentRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentRemoveCommand_instances, "m", _EntraPimRoleAssignmentRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const { userId, userName, groupId, groupName, ticketNumber, ticketSystem } = args.options;
        try {
            const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
            const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);
            if (isAppOnlyAccessToken && !userId && !userName && !groupId && !groupName) {
                throw 'When running with application permissions either userId, userName, groupId or groupName is required';
            }
            const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
            const principalId = await this.getPrincipalId(args.options, logger);
            const requestOptions = {
                url: `${this.resource}/v1.0/roleManagement/directory/roleAssignmentScheduleRequests`,
                headers: {
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    principalId: principalId,
                    roleDefinitionId: roleDefinitionId,
                    directoryScopeId: this.getDirectoryScope(args.options),
                    action: !userId && !userName && !groupId && !groupName ? 'selfDeactivate' : 'adminRemove',
                    justification: args.options.justification,
                    ticketInfo: {
                        ticketNumber: ticketNumber,
                        ticketSystem: ticketSystem
                    }
                }
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getRoleDefinitionId(options, logger) {
        if (options.roleDefinitionId) {
            return options.roleDefinitionId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving role definition by its name '${options.roleDefinitionName}'`);
        }
        const role = await roleDefinition.getRoleDefinitionByDisplayName(options.roleDefinitionName);
        return role.id;
    }
    async getPrincipalId(options, logger) {
        if (options.userId || options.groupId) {
            return options.userId || options.groupId;
        }
        if (options.userName) {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving user by its name '${options.userName}'`);
            }
            return await entraUser.getUserIdByUpn(options.userName);
        }
        else if (options.groupName) {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving group by its name '${options.groupName}'`);
            }
            return await entraGroup.getGroupIdByDisplayName(options.groupName);
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving id of the current user`);
        }
        const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
        return accessToken.getUserIdFromAccessToken(token);
    }
    getDirectoryScope(options) {
        if (options.administrativeUnitId) {
            return `/administrativeUnits/${options.administrativeUnitId}`;
        }
        if (options.applicationId) {
            return `/${options.applicationId}`;
        }
        return '/';
    }
}
_EntraPimRoleAssignmentRemoveCommand_instances = new WeakSet(), _EntraPimRoleAssignmentRemoveCommand_initTelemetry = function _EntraPimRoleAssignmentRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            roleDefinitionName: typeof args.options.roleDefinitionName !== 'undefined',
            roleDefinitionId: typeof args.options.roleDefinitionId !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            administrativeUnitId: typeof args.options.administrativeUnitId !== 'undefined',
            applicationId: typeof args.options.applicationId !== 'undefined',
            justification: typeof args.options.justification !== 'undefined',
            ticketNumber: typeof args.options.ticketNumber !== 'undefined',
            ticketSystem: typeof args.options.ticketSystem !== 'undefined'
        });
    });
}, _EntraPimRoleAssignmentRemoveCommand_initOptions = function _EntraPimRoleAssignmentRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-n, --roleDefinitionName [roleDefinitionName]'
    }, {
        option: '-i, --roleDefinitionId [roleDefinitionId]'
    }, {
        option: "--userId [userId]"
    }, {
        option: "--userName [userName]"
    }, {
        option: "--groupId [groupId]"
    }, {
        option: "--groupName [groupName]"
    }, {
        option: "--administrativeUnitId [administrativeUnitId]"
    }, {
        option: "--applicationId [applicationId]"
    }, {
        option: "-j, --justification [justification]"
    }, {
        option: "--ticketNumber [ticketNumber]"
    }, {
        option: "--ticketSystem [ticketSystem]"
    });
}, _EntraPimRoleAssignmentRemoveCommand_initValidators = function _EntraPimRoleAssignmentRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.roleDefinitionId && !validation.isValidGuid(args.options.roleDefinitionId)) {
            return `${args.options.roleDefinitionId} is not a valid GUID`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID`;
        }
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId)) {
            return `${args.options.administrativeUnitId} is not a valid GUID`;
        }
        if (args.options.applicationId && !validation.isValidGuid(args.options.applicationId)) {
            return `${args.options.applicationId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraPimRoleAssignmentRemoveCommand_initOptionSets = function _EntraPimRoleAssignmentRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['roleDefinitionName', 'roleDefinitionId'] });
    this.optionSets.push({
        options: ['userId', 'userName', 'groupId', 'groupName'],
        runsWhen: (args) => {
            return args.options.userId !== undefined || args.options.userName !== undefined || args.options.groupId !== undefined || args.options.groupName !== undefined;
        }
    });
    this.optionSets.push({
        options: ['administrativeUnitId', 'applicationId'],
        runsWhen: (args) => {
            return args.options.administrativeUnitId !== undefined || args.options.applicationId !== undefined;
        }
    });
}, _EntraPimRoleAssignmentRemoveCommand_initTypes = function _EntraPimRoleAssignmentRemoveCommand_initTypes() {
    this.types.string.push('userId', 'userName', 'groupId', 'groupName', 'administrativeUnitId', 'applicationId', 'roleDefinitionName', 'roleDefinitionId', 'justification', 'ticketNumber', 'ticketSystem');
};
export default new EntraPimRoleAssignmentRemoveCommand();
//# sourceMappingURL=pim-role-assignment-remove.js.map