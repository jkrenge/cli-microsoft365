var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskReferenceListCommand_instances, _PlannerTaskReferenceListCommand_initOptions, _PlannerTaskReferenceListCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerTaskReferenceListCommand extends GraphCommand {
    get name() {
        return commands.TASK_REFERENCE_LIST;
    }
    get description() {
        return 'Retrieve the references of the specified planner task';
    }
    constructor() {
        super();
        _PlannerTaskReferenceListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceListCommand_instances, "m", _PlannerTaskReferenceListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceListCommand_instances, "m", _PlannerTaskReferenceListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(args.options.taskId)}/details?$select=references`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res.references);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PlannerTaskReferenceListCommand_instances = new WeakSet(), _PlannerTaskReferenceListCommand_initOptions = function _PlannerTaskReferenceListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --taskId <taskId>'
    });
}, _PlannerTaskReferenceListCommand_initTypes = function _PlannerTaskReferenceListCommand_initTypes() {
    this.types.string.push('taskId');
};
export default new PlannerTaskReferenceListCommand();
//# sourceMappingURL=task-reference-list.js.map