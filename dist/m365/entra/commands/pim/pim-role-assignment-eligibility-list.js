var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraPimRoleAssignmentEligibilityListCommand_instances, _EntraPimRoleAssignmentEligibilityListCommand_initTelemetry, _EntraPimRoleAssignmentEligibilityListCommand_initOptions, _EntraPimRoleAssignmentEligibilityListCommand_initValidators, _EntraPimRoleAssignmentEligibilityListCommand_initOptionSets;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
class EntraPimRoleAssignmentEligibilityListCommand extends GraphCommand {
    get name() {
        return commands.PIM_ROLE_ASSIGNMENT_ELIGIBILITY_LIST;
    }
    get description() {
        return 'Retrieves a list of eligible roles a user or group can be assigned to';
    }
    defaultProperties() {
        return ['roleDefinitionId', 'roleDefinitionName', 'principalId'];
    }
    constructor() {
        super();
        _EntraPimRoleAssignmentEligibilityListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentEligibilityListCommand_instances, "m", _EntraPimRoleAssignmentEligibilityListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentEligibilityListCommand_instances, "m", _EntraPimRoleAssignmentEligibilityListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentEligibilityListCommand_instances, "m", _EntraPimRoleAssignmentEligibilityListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentEligibilityListCommand_instances, "m", _EntraPimRoleAssignmentEligibilityListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of eligible roles for ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName || 'all users'}...`);
        }
        const queryParameters = [];
        const expands = [];
        try {
            const principalId = await this.getPrincipalId(logger, args.options);
            if (principalId) {
                queryParameters.push(`$filter=principalId eq '${principalId}'`);
            }
            expands.push('roleDefinition($select=displayName)');
            if (args.options.withPrincipalDetails) {
                expands.push('principal');
            }
            queryParameters.push(`$expand=${expands.join(',')}`);
            const url = `${this.resource}/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?${queryParameters.join('&')}`;
            const results = await odata.getAllItems(url);
            results.forEach(c => {
                const roleDefinition = c['roleDefinition'];
                if (roleDefinition) {
                    c.roleDefinitionName = roleDefinition.displayName;
                }
                delete c['roleDefinition'];
            });
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getPrincipalId(logger, options) {
        let principalId = options.userId;
        if (options.userName) {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving user by its name '${options.userName}'`);
            }
            principalId = await entraUser.getUserIdByUpn(options.userName);
        }
        else if (options.groupId) {
            principalId = options.groupId;
        }
        else if (options.groupName) {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving group by its name '${options.groupName}'`);
            }
            principalId = await entraGroup.getGroupIdByDisplayName(options.groupName);
        }
        return principalId;
    }
}
_EntraPimRoleAssignmentEligibilityListCommand_instances = new WeakSet(), _EntraPimRoleAssignmentEligibilityListCommand_initTelemetry = function _EntraPimRoleAssignmentEligibilityListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            withPrincipalDetails: !!args.options.withPrincipalDetails
        });
    });
}, _EntraPimRoleAssignmentEligibilityListCommand_initOptions = function _EntraPimRoleAssignmentEligibilityListCommand_initOptions() {
    this.options.unshift({
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '--withPrincipalDetails'
    });
}, _EntraPimRoleAssignmentEligibilityListCommand_initValidators = function _EntraPimRoleAssignmentEligibilityListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `'${args.options.userId} is not a valid GUID for option 'userId'.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `'${args.options.userName} is not a valid user principal name for option 'userName'.`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `'${args.options.groupId}' is not a valid GUID for option 'groupId'.`;
        }
        return true;
    });
}, _EntraPimRoleAssignmentEligibilityListCommand_initOptionSets = function _EntraPimRoleAssignmentEligibilityListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName', 'groupId', 'groupName'],
        runsWhen: (args) => args.options.userId || args.options.userName || args.options.groupName || args.options.groupId
    });
};
export default new EntraPimRoleAssignmentEligibilityListCommand();
//# sourceMappingURL=pim-role-assignment-eligibility-list.js.map