var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskChecklistItemAddCommand_instances, _PlannerTaskChecklistItemAddCommand_initTelemetry, _PlannerTaskChecklistItemAddCommand_initOptions, _PlannerTaskChecklistItemAddCommand_initTypes;
import { v4 } from 'uuid';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerTaskChecklistItemAddCommand extends GraphCommand {
    get name() {
        return commands.TASK_CHECKLISTITEM_ADD;
    }
    get description() {
        return 'Adds a new checklist item to a Planner task.';
    }
    constructor() {
        super();
        _PlannerTaskChecklistItemAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTaskChecklistItemAddCommand_instances, "m", _PlannerTaskChecklistItemAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTaskChecklistItemAddCommand_instances, "m", _PlannerTaskChecklistItemAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskChecklistItemAddCommand_instances, "m", _PlannerTaskChecklistItemAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const etag = await this.getTaskDetailsEtag(args.options.taskId);
            const body = {
                checklist: {
                    // Generate new GUID for new task checklist item
                    [v4()]: {
                        '@odata.type': 'microsoft.graph.plannerChecklistItem',
                        title: args.options.title,
                        isChecked: args.options.isChecked || false
                    }
                }
            };
            const requestOptions = {
                url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(args.options.taskId)}/details`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    prefer: 'return=representation',
                    'if-match': etag
                },
                responseType: 'json',
                data: body
            };
            const result = await request.patch(requestOptions);
            if (!cli.shouldTrimOutput(args.options.output)) {
                await logger.log(result.checklist);
            }
            else {
                // Transform checklist item object to text friendly format
                const output = Object.getOwnPropertyNames(result.checklist).map(prop => ({ id: prop, ...result.checklist[prop] }));
                await logger.log(output);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTaskDetailsEtag(taskId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details`,
            headers: {
                accept: 'application/json;odata.metadata=minimal'
            },
            responseType: 'json'
        };
        const task = await request.get(requestOptions);
        return task['@odata.etag'];
    }
}
_PlannerTaskChecklistItemAddCommand_instances = new WeakSet(), _PlannerTaskChecklistItemAddCommand_initTelemetry = function _PlannerTaskChecklistItemAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            isChecked: !!args.options.isChecked
        });
    });
}, _PlannerTaskChecklistItemAddCommand_initOptions = function _PlannerTaskChecklistItemAddCommand_initOptions() {
    this.options.unshift({ option: '-i, --taskId <taskId>' }, { option: '-t, --title <title>' }, { option: '--isChecked' });
}, _PlannerTaskChecklistItemAddCommand_initTypes = function _PlannerTaskChecklistItemAddCommand_initTypes() {
    this.types.string.push('title', 'taskId');
};
export default new PlannerTaskChecklistItemAddCommand();
//# sourceMappingURL=task-checklistitem-add.js.map