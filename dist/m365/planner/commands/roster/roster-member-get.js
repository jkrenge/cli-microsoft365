var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerRosterMemberGetCommand_instances, _PlannerRosterMemberGetCommand_initTelemetry, _PlannerRosterMemberGetCommand_initOptions, _PlannerRosterMemberGetCommand_initOptionSets, _PlannerRosterMemberGetCommand_initValidators, _PlannerRosterMemberGetCommand_initTypes;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';
class PlannerRosterMemberGetCommand extends GraphCommand {
    get name() {
        return commands.ROSTER_MEMBER_GET;
    }
    get description() {
        return 'Gets a member of the specified Microsoft Planner Roster';
    }
    constructor() {
        super();
        _PlannerRosterMemberGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberGetCommand_instances, "m", _PlannerRosterMemberGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberGetCommand_instances, "m", _PlannerRosterMemberGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberGetCommand_instances, "m", _PlannerRosterMemberGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberGetCommand_instances, "m", _PlannerRosterMemberGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberGetCommand_instances, "m", _PlannerRosterMemberGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving member ${args.options.userName || args.options.userId} from the Microsoft Planner Roster`);
        }
        try {
            const userId = await this.getUserId(args);
            const requestOptions = {
                url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members/${userId}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const response = await request.get(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUserId(args) {
        if (args.options.userId) {
            return args.options.userId;
        }
        return entraUser.getUserIdByUpn(args.options.userName);
    }
}
_PlannerRosterMemberGetCommand_instances = new WeakSet(), _PlannerRosterMemberGetCommand_initTelemetry = function _PlannerRosterMemberGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _PlannerRosterMemberGetCommand_initOptions = function _PlannerRosterMemberGetCommand_initOptions() {
    this.options.unshift({
        option: '--rosterId <rosterId>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _PlannerRosterMemberGetCommand_initOptionSets = function _PlannerRosterMemberGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName'] });
}, _PlannerRosterMemberGetCommand_initValidators = function _PlannerRosterMemberGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        return true;
    });
}, _PlannerRosterMemberGetCommand_initTypes = function _PlannerRosterMemberGetCommand_initTypes() {
    this.types.string.push('rosterId', 'userId', 'userName');
};
export default new PlannerRosterMemberGetCommand();
//# sourceMappingURL=roster-member-get.js.map