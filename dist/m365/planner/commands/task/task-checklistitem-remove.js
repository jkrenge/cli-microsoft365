var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskChecklistItemRemoveCommand_instances, _PlannerTaskChecklistItemRemoveCommand_initTelemetry, _PlannerTaskChecklistItemRemoveCommand_initOptions, _PlannerTaskChecklistItemRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerTaskChecklistItemRemoveCommand extends GraphCommand {
    get name() {
        return commands.TASK_CHECKLISTITEM_REMOVE;
    }
    get description() {
        return 'Removes the checklist item from the planner task';
    }
    constructor() {
        super();
        _PlannerTaskChecklistItemRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTaskChecklistItemRemoveCommand_instances, "m", _PlannerTaskChecklistItemRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTaskChecklistItemRemoveCommand_instances, "m", _PlannerTaskChecklistItemRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskChecklistItemRemoveCommand_instances, "m", _PlannerTaskChecklistItemRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeChecklistitem(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the checklist item with id ${args.options.id} from the planner task?` });
            if (result) {
                await this.removeChecklistitem(args);
            }
        }
    }
    async removeChecklistitem(args) {
        try {
            const task = await this.getTaskDetails(args.options.taskId);
            if (!task.checklist || !task.checklist[args.options.id]) {
                throw `The specified checklist item with id ${args.options.id} does not exist`;
            }
            const requestOptionsTaskDetails = {
                url: `${this.resource}/v1.0/planner/tasks/${args.options.taskId}/details`,
                headers: {
                    'accept': 'application/json;odata.metadata=none',
                    'If-Match': task['@odata.etag'],
                    'Prefer': 'return=representation'
                },
                responseType: 'json',
                data: {
                    checklist: {
                        [args.options.id]: null
                    }
                }
            };
            await request.patch(requestOptionsTaskDetails);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTaskDetails(taskId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details?$select=checklist`,
            headers: {
                accept: 'application/json;odata.metadata=minimal'
            },
            responseType: 'json'
        };
        return await request.get(requestOptions);
    }
}
_PlannerTaskChecklistItemRemoveCommand_instances = new WeakSet(), _PlannerTaskChecklistItemRemoveCommand_initTelemetry = function _PlannerTaskChecklistItemRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _PlannerTaskChecklistItemRemoveCommand_initOptions = function _PlannerTaskChecklistItemRemoveCommand_initOptions() {
    this.options.unshift({ option: '-i, --id <id>' }, { option: '--taskId <taskId>' }, { option: '-f, --force' });
}, _PlannerTaskChecklistItemRemoveCommand_initTypes = function _PlannerTaskChecklistItemRemoveCommand_initTypes() {
    this.types.string.push('id', 'taskId');
};
export default new PlannerTaskChecklistItemRemoveCommand();
//# sourceMappingURL=task-checklistitem-remove.js.map