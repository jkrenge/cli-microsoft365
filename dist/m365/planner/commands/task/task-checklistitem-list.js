var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskChecklistItemListCommand_instances, _PlannerTaskChecklistItemListCommand_initOptions, _PlannerTaskChecklistItemListCommand_initTypes;
import { cli } from "../../../../cli/cli.js";
import request from "../../../../request.js";
import { formatting } from "../../../../utils/formatting.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
class PlannerTaskChecklistItemListCommand extends GraphCommand {
    get name() {
        return commands.TASK_CHECKLISTITEM_LIST;
    }
    get description() {
        return 'Lists the checklist items of a Planner task.';
    }
    defaultProperties() {
        return ['id', 'title', 'isChecked'];
    }
    constructor() {
        super();
        _PlannerTaskChecklistItemListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTaskChecklistItemListCommand_instances, "m", _PlannerTaskChecklistItemListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskChecklistItemListCommand_instances, "m", _PlannerTaskChecklistItemListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(args.options.taskId)}/details?$select=checklist`,
            headers: {
                accept: "application/json;odata.metadata=none"
            },
            responseType: "json"
        };
        try {
            const res = await request.get(requestOptions);
            if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
                await logger.log(res.checklist);
            }
            else {
                //converted to text friendly output
                const output = Object.getOwnPropertyNames(res.checklist).map(prop => ({ id: prop, ...res.checklist[prop] }));
                await logger.log(output);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PlannerTaskChecklistItemListCommand_instances = new WeakSet(), _PlannerTaskChecklistItemListCommand_initOptions = function _PlannerTaskChecklistItemListCommand_initOptions() {
    this.options.unshift({
        option: "-i, --taskId <taskId>"
    });
}, _PlannerTaskChecklistItemListCommand_initTypes = function _PlannerTaskChecklistItemListCommand_initTypes() {
    this.types.string.push('taskId');
};
export default new PlannerTaskChecklistItemListCommand();
//# sourceMappingURL=task-checklistitem-list.js.map