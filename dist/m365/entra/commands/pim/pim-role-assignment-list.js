var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraPimRoleAssignmentListCommand_instances, _EntraPimRoleAssignmentListCommand_initTelemetry, _EntraPimRoleAssignmentListCommand_initOptions, _EntraPimRoleAssignmentListCommand_initValidators, _EntraPimRoleAssignmentListCommand_initOptionSets;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
class EntraPimRoleAssignmentListCommand extends GraphCommand {
    get name() {
        return commands.PIM_ROLE_ASSIGNMENT_LIST;
    }
    get description() {
        return 'Retrieves a list of Entra role assignments for a user or group';
    }
    constructor() {
        super();
        _EntraPimRoleAssignmentListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentListCommand_instances, "m", _EntraPimRoleAssignmentListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentListCommand_instances, "m", _EntraPimRoleAssignmentListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentListCommand_instances, "m", _EntraPimRoleAssignmentListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleAssignmentListCommand_instances, "m", _EntraPimRoleAssignmentListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const queryParameters = [];
        const filters = [];
        const expands = [];
        try {
            const principalId = await this.getPrincipalId(logger, args.options);
            if (principalId) {
                filters.push(`principalId eq '${principalId}'`);
            }
            if (args.options.startDateTime) {
                filters.push(`startDateTime ge ${args.options.startDateTime}`);
            }
            if (filters.length > 0) {
                queryParameters.push(`$filter=${filters.join(' and ')}`);
            }
            expands.push('roleDefinition($select=displayName)');
            if (args.options.withPrincipalDetails) {
                expands.push('principal');
            }
            queryParameters.push(`$expand=${expands.join(',')}`);
            const queryString = `?${queryParameters.join('&')}`;
            const url = `${this.resource}/v1.0/roleManagement/directory/roleAssignmentScheduleInstances${queryString}`;
            const results = await odata.getAllItems(url);
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
_EntraPimRoleAssignmentListCommand_instances = new WeakSet(), _EntraPimRoleAssignmentListCommand_initTelemetry = function _EntraPimRoleAssignmentListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            startDateTime: typeof args.options.startDateTime !== 'undefined',
            withPrincipalDetails: !!args.options.withPrincipalDetails
        });
    });
}, _EntraPimRoleAssignmentListCommand_initOptions = function _EntraPimRoleAssignmentListCommand_initOptions() {
    this.options.unshift({
        option: "--userId [userId]"
    }, {
        option: "--userName [userName]"
    }, {
        option: "--groupId [groupId]"
    }, {
        option: "--groupName [groupName]"
    }, {
        option: "-s, --startDateTime [startDateTime]"
    }, {
        option: "--withPrincipalDetails"
    });
}, _EntraPimRoleAssignmentListCommand_initValidators = function _EntraPimRoleAssignmentListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID`;
        }
        if (args.options.startDateTime && !validation.isValidISODateTime(args.options.startDateTime)) {
            return `${args.options.startDateTime} is not a valid ISO 8601 date time string`;
        }
        return true;
    });
}, _EntraPimRoleAssignmentListCommand_initOptionSets = function _EntraPimRoleAssignmentListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName', 'groupId', 'groupName'],
        runsWhen: (args) => args.options.userId || args.options.userName || args.options.groupId || args.options.groupName
    });
};
export default new EntraPimRoleAssignmentListCommand();
//# sourceMappingURL=pim-role-assignment-list.js.map