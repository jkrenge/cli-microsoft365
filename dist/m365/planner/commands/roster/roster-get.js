var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerRosterGetCommand_instances, _PlannerRosterGetCommand_initOptions, _PlannerRosterGetCommand_initTypes;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerRosterGetCommand extends GraphCommand {
    get name() {
        return commands.ROSTER_GET;
    }
    get description() {
        return 'Retrieve information about a specific Microsoft Planner Roster';
    }
    constructor() {
        super();
        _PlannerRosterGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerRosterGetCommand_instances, "m", _PlannerRosterGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerRosterGetCommand_instances, "m", _PlannerRosterGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about Microsoft Planner Roster with id ${args.options.id}`);
        }
        const requestOptions = {
            url: `${this.resource}/beta/planner/rosters/${args.options.id}`,
            headers: {
                'accept': 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const response = await request.get(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PlannerRosterGetCommand_instances = new WeakSet(), _PlannerRosterGetCommand_initOptions = function _PlannerRosterGetCommand_initOptions() {
    this.options.unshift({
        option: '--id <id>'
    });
}, _PlannerRosterGetCommand_initTypes = function _PlannerRosterGetCommand_initTypes() {
    this.types.string.push('id');
};
export default new PlannerRosterGetCommand();
//# sourceMappingURL=roster-get.js.map