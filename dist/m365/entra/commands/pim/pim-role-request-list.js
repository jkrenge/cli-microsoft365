var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraPimRoleRequestListCommand_instances, _EntraPimRoleRequestListCommand_initTelemetry, _EntraPimRoleRequestListCommand_initOptions, _EntraPimRoleRequestListCommand_initValidators, _EntraPimRoleRequestListCommand_initOptionSets;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
class EntraPimRoleRequestListCommand extends GraphCommand {
    get name() {
        return commands.PIM_ROLE_REQUEST_LIST;
    }
    get description() {
        return 'Retrieves a list of PIM requests for roles';
    }
    defaultProperties() {
        return ['id', 'roleDefinitionName', 'principalId'];
    }
    constructor() {
        super();
        _EntraPimRoleRequestListCommand_instances.add(this);
        this.allowedStatuses = ['Canceled', 'Denied', 'Failed', 'Granted', 'PendingAdminDecision', 'PendingApproval', 'PendingProvisioning', 'PendingScheduleCreation', 'Provisioned', 'Revoked', 'ScheduleCreated'];
        __classPrivateFieldGet(this, _EntraPimRoleRequestListCommand_instances, "m", _EntraPimRoleRequestListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleRequestListCommand_instances, "m", _EntraPimRoleRequestListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleRequestListCommand_instances, "m", _EntraPimRoleRequestListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraPimRoleRequestListCommand_instances, "m", _EntraPimRoleRequestListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of PIM roles requests for ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName || 'all users'}...`);
        }
        const queryParameters = [];
        const filters = [];
        const expands = [];
        try {
            const principalId = await this.getPrincipalId(logger, args.options);
            if (principalId) {
                filters.push(`principalId eq '${principalId}'`);
            }
            if (args.options.createdDateTime) {
                filters.push(`createdDateTime ge ${args.options.createdDateTime}`);
            }
            if (args.options.status) {
                filters.push(`status eq '${args.options.status}'`);
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
            const url = `${this.resource}/v1.0/roleManagement/directory/roleAssignmentScheduleRequests${queryString}`;
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
_EntraPimRoleRequestListCommand_instances = new WeakSet(), _EntraPimRoleRequestListCommand_initTelemetry = function _EntraPimRoleRequestListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            createdDateTime: typeof args.options.createdDateTime !== 'undefined',
            status: typeof args.options.status !== 'undefined',
            withPrincipalDetails: !!args.options.withPrincipalDetails
        });
    });
}, _EntraPimRoleRequestListCommand_initOptions = function _EntraPimRoleRequestListCommand_initOptions() {
    this.options.unshift({
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '-c, --createdDateTime [createdDateTime]'
    }, {
        option: '-s, --status [status]',
        autocomplete: this.allowedStatuses
    }, {
        option: '--withPrincipalDetails'
    });
}, _EntraPimRoleRequestListCommand_initValidators = function _EntraPimRoleRequestListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `'${args.options.userId}' is not a valid GUID for option 'userId'`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `'${args.options.userName}' is not a valid user principal name for option 'userName'.`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `'${args.options.groupId}' is not a valid GUID for option 'groupId'`;
        }
        if (args.options.createdDateTime && !validation.isValidISODateTime(args.options.createdDateTime)) {
            return `'${args.options.createdDateTime}' is not a valid ISO 8601 date time string for option 'createdDateTime'`;
        }
        if (args.options.status && !this.allowedStatuses.some(status => status.toLowerCase() === args.options.status.toLowerCase())) {
            return `'${args.options.status}' for option 'status' must be one of the following values: ${this.allowedStatuses.join(', ')}.`;
        }
        return true;
    });
}, _EntraPimRoleRequestListCommand_initOptionSets = function _EntraPimRoleRequestListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName', 'groupId', 'groupName'],
        runsWhen: (args) => args.options.userId || args.options.userName || args.options.groupId || args.options.groupName
    });
};
export default new EntraPimRoleRequestListCommand();
//# sourceMappingURL=pim-role-request-list.js.map