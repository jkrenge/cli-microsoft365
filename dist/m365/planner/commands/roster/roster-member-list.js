var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerRosterMemberListCommand_instances, _PlannerRosterMemberListCommand_initOptions, _PlannerRosterMemberListCommand_initTypes;
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerRosterMemberListCommand extends GraphCommand {
    get name() {
        return commands.ROSTER_MEMBER_LIST;
    }
    get description() {
        return 'Lists members of the specified Microsoft Planner Roster';
    }
    constructor() {
        super();
        _PlannerRosterMemberListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberListCommand_instances, "m", _PlannerRosterMemberListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberListCommand_instances, "m", _PlannerRosterMemberListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr('Retrieving members of the specified Microsoft Planner Roster');
        }
        try {
            const response = await odata.getAllItems(`${this.resource}/beta/planner/rosters/${args.options.rosterId}/members`);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PlannerRosterMemberListCommand_instances = new WeakSet(), _PlannerRosterMemberListCommand_initOptions = function _PlannerRosterMemberListCommand_initOptions() {
    this.options.unshift({
        option: '--rosterId <rosterId>'
    });
}, _PlannerRosterMemberListCommand_initTypes = function _PlannerRosterMemberListCommand_initTypes() {
    this.types.string.push('rosterId');
};
export default new PlannerRosterMemberListCommand();
//# sourceMappingURL=roster-member-list.js.map