var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAdministrativeUnitRoleAssignmentAddCommand_instances, _EntraAdministrativeUnitRoleAssignmentAddCommand_initTelemetry, _EntraAdministrativeUnitRoleAssignmentAddCommand_initOptions, _EntraAdministrativeUnitRoleAssignmentAddCommand_initValidators, _EntraAdministrativeUnitRoleAssignmentAddCommand_initOptionSets;
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { roleAssignment } from '../../../../utils/roleAssignment.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraAdministrativeUnitRoleAssignmentAddCommand extends GraphCommand {
    get name() {
        return commands.ADMINISTRATIVEUNIT_ROLEASSIGNMENT_ADD;
    }
    get description() {
        return 'Assigns a Microsoft Entra role with administrative unit scope to a user';
    }
    constructor() {
        super();
        _EntraAdministrativeUnitRoleAssignmentAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitRoleAssignmentAddCommand_instances, "m", _EntraAdministrativeUnitRoleAssignmentAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitRoleAssignmentAddCommand_instances, "m", _EntraAdministrativeUnitRoleAssignmentAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitRoleAssignmentAddCommand_instances, "m", _EntraAdministrativeUnitRoleAssignmentAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAdministrativeUnitRoleAssignmentAddCommand_instances, "m", _EntraAdministrativeUnitRoleAssignmentAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            let { administrativeUnitId, roleDefinitionId, userId } = args.options;
            if (args.options.administrativeUnitName) {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving administrative unit by its name '${args.options.administrativeUnitName}'`);
                }
                administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id;
            }
            if (args.options.roleDefinitionName) {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving role definition by its name '${args.options.roleDefinitionName}'`);
                }
                roleDefinitionId = (await roleDefinition.getRoleDefinitionByDisplayName(args.options.roleDefinitionName)).id;
            }
            if (args.options.userName) {
                if (this.verbose) {
                    await logger.logToStderr(`Retrieving user by UPN '${args.options.userName}'`);
                }
                userId = await entraUser.getUserIdByUpn(args.options.userName);
            }
            const unifiedRoleAssignment = await roleAssignment.createRoleAssignmentWithAdministrativeUnitScope(roleDefinitionId, userId, administrativeUnitId);
            await logger.log(unifiedRoleAssignment);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraAdministrativeUnitRoleAssignmentAddCommand_instances = new WeakSet(), _EntraAdministrativeUnitRoleAssignmentAddCommand_initTelemetry = function _EntraAdministrativeUnitRoleAssignmentAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            administrativeUnitId: typeof args.options.administrativeUnitId !== 'undefined',
            administrativeUnitName: typeof args.options.administrativeUnitName !== 'undefined',
            roleDefinitionId: typeof args.options.roleDefinitionId !== 'undefined',
            roleDefinitionName: typeof args.options.roleDefinitionName !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _EntraAdministrativeUnitRoleAssignmentAddCommand_initOptions = function _EntraAdministrativeUnitRoleAssignmentAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --administrativeUnitId [administrativeUnitId]'
    }, {
        option: '-n, --administrativeUnitName [administrativeUnitName]'
    }, {
        option: '--roleDefinitionId [roleDefinitionId]'
    }, {
        option: '--roleDefinitionName [roleDefinitionName]'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _EntraAdministrativeUnitRoleAssignmentAddCommand_initValidators = function _EntraAdministrativeUnitRoleAssignmentAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId)) {
            return `${args.options.administrativeUnitId} is not a valid GUID`;
        }
        if (args.options.roleDefinitionId && !validation.isValidGuid(args.options.roleDefinitionId)) {
            return `${args.options.roleDefinitionId} is not a valid GUID`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraAdministrativeUnitRoleAssignmentAddCommand_initOptionSets = function _EntraAdministrativeUnitRoleAssignmentAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['administrativeUnitId', 'administrativeUnitName'] });
    this.optionSets.push({ options: ['roleDefinitionId', 'roleDefinitionName'] });
    this.optionSets.push({ options: ['userId', 'userName'] });
};
export default new EntraAdministrativeUnitRoleAssignmentAddCommand();
//# sourceMappingURL=administrativeunit-roleassignment-add.js.map