var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerRosterPlanListCommand_instances, _PlannerRosterPlanListCommand_initTelemetry, _PlannerRosterPlanListCommand_initOptions, _PlannerRosterPlanListCommand_initValidators, _PlannerRosterPlanListCommand_initTypes, _PlannerRosterPlanListCommand_initOptionSets;
import auth from '../../../../Auth.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
class PlannerRosterPlanListCommand extends GraphCommand {
    get name() {
        return commands.ROSTER_PLAN_LIST;
    }
    get description() {
        return 'Lists all Microsoft Planner Roster plans for a specific user';
    }
    constructor() {
        super();
        _PlannerRosterPlanListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerRosterPlanListCommand_instances, "m", _PlannerRosterPlanListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerRosterPlanListCommand_instances, "m", _PlannerRosterPlanListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerRosterPlanListCommand_instances, "m", _PlannerRosterPlanListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerRosterPlanListCommand_instances, "m", _PlannerRosterPlanListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerRosterPlanListCommand_instances, "m", _PlannerRosterPlanListCommand_initTypes).call(this);
    }
    defaultProperties() {
        return ['id', 'title', 'createdDateTime', 'owner'];
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
        if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
            this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
        }
        else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName)) {
            this.handleError(`The options 'userId' or 'userName' cannot be used when obtaining Microsoft Planner Roster plans using delegated permissions.`);
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving Microsoft Planner Roster plans for user: ${args.options.userId || args.options.userName || 'current user'}.`);
        }
        let requestUrl = `${this.resource}/beta/`;
        if (args.options.userId || args.options.userName) {
            requestUrl += `users/${args.options.userId || args.options.userName}`;
        }
        else {
            requestUrl += 'me';
        }
        requestUrl += `/planner/rosterPlans`;
        try {
            const items = await odata.getAllItems(requestUrl);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PlannerRosterPlanListCommand_instances = new WeakSet(), _PlannerRosterPlanListCommand_initTelemetry = function _PlannerRosterPlanListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _PlannerRosterPlanListCommand_initOptions = function _PlannerRosterPlanListCommand_initOptions() {
    this.options.unshift({
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _PlannerRosterPlanListCommand_initValidators = function _PlannerRosterPlanListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN)`;
        }
        return true;
    });
}, _PlannerRosterPlanListCommand_initTypes = function _PlannerRosterPlanListCommand_initTypes() {
    this.types.string.push('userId', 'userName');
}, _PlannerRosterPlanListCommand_initOptionSets = function _PlannerRosterPlanListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName'],
        runsWhen: (args) => args.options.userId || args.options.userName
    });
};
export default new PlannerRosterPlanListCommand();
//# sourceMappingURL=roster-plan-list.js.map