var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerRosterMemberAddCommand_instances, _PlannerRosterMemberAddCommand_initTelemetry, _PlannerRosterMemberAddCommand_initOptions, _PlannerRosterMemberAddCommand_initOptionSets, _PlannerRosterMemberAddCommand_initValidators, _PlannerRosterMemberAddCommand_initTypes;
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerRosterMemberAddCommand extends GraphCommand {
    get name() {
        return commands.ROSTER_MEMBER_ADD;
    }
    get description() {
        return 'Adds a user to a Microsoft Planner Roster';
    }
    constructor() {
        super();
        _PlannerRosterMemberAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberAddCommand_instances, "m", _PlannerRosterMemberAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberAddCommand_instances, "m", _PlannerRosterMemberAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberAddCommand_instances, "m", _PlannerRosterMemberAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberAddCommand_instances, "m", _PlannerRosterMemberAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberAddCommand_instances, "m", _PlannerRosterMemberAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Adding a user to a Microsoft Planner Roster');
        }
        try {
            const userId = await this.getUserId(logger, args);
            const requestOptions = {
                url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                data: {
                    userId: userId
                },
                responseType: 'json'
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUserId(logger, args) {
        if (this.verbose) {
            await logger.logToStderr("Getting the user ID");
        }
        if (args.options.userId) {
            return args.options.userId;
        }
        const userId = await entraUser.getUserIdByUpn(args.options.userName);
        return userId;
    }
}
_PlannerRosterMemberAddCommand_instances = new WeakSet(), _PlannerRosterMemberAddCommand_initTelemetry = function _PlannerRosterMemberAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _PlannerRosterMemberAddCommand_initOptions = function _PlannerRosterMemberAddCommand_initOptions() {
    this.options.unshift({
        option: '--rosterId <rosterId>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _PlannerRosterMemberAddCommand_initOptionSets = function _PlannerRosterMemberAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName'] });
}, _PlannerRosterMemberAddCommand_initValidators = function _PlannerRosterMemberAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN)`;
        }
        return true;
    });
}, _PlannerRosterMemberAddCommand_initTypes = function _PlannerRosterMemberAddCommand_initTypes() {
    this.types.string.push('rosterId', 'userId', 'userName');
};
export default new PlannerRosterMemberAddCommand();
//# sourceMappingURL=roster-member-add.js.map