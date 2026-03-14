var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskReferenceAddCommand_instances, _PlannerTaskReferenceAddCommand_initTelemetry, _PlannerTaskReferenceAddCommand_initOptions, _PlannerTaskReferenceAddCommand_initValidators, _PlannerTaskReferenceAddCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerTaskReferenceAddCommand extends GraphCommand {
    get name() {
        return commands.TASK_REFERENCE_ADD;
    }
    get description() {
        return 'Adds a new reference to a Planner task';
    }
    constructor() {
        super();
        _PlannerTaskReferenceAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceAddCommand_instances, "m", _PlannerTaskReferenceAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceAddCommand_instances, "m", _PlannerTaskReferenceAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceAddCommand_instances, "m", _PlannerTaskReferenceAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerTaskReferenceAddCommand_instances, "m", _PlannerTaskReferenceAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const etag = await this.getTaskDetailsEtag(args.options.taskId);
            const requestOptionsTaskDetails = {
                url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(args.options.taskId)}/details`,
                headers: {
                    'accept': 'application/json;odata.metadata=none',
                    'If-Match': etag,
                    'Prefer': 'return=representation'
                },
                responseType: 'json',
                data: {
                    references: {
                        [formatting.openTypesEncoder(args.options.url)]: {
                            '@odata.type': 'microsoft.graph.plannerExternalReference',
                            previewPriority: ' !',
                            ...(args.options.alias && { alias: args.options.alias }),
                            ...(args.options.type && { type: args.options.type })
                        }
                    }
                }
            };
            const res = await request.patch(requestOptionsTaskDetails);
            await logger.log(res.references);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTaskDetailsEtag(taskId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}/details`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response['@odata.etag'];
    }
}
_PlannerTaskReferenceAddCommand_instances = new WeakSet(), _PlannerTaskReferenceAddCommand_initTelemetry = function _PlannerTaskReferenceAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            alias: typeof args.options.alias !== 'undefined',
            type: args.options.type
        });
    });
}, _PlannerTaskReferenceAddCommand_initOptions = function _PlannerTaskReferenceAddCommand_initOptions() {
    this.options.unshift({ option: '-i, --taskId <taskId>' }, { option: '-u, --url <url>' }, { option: '--alias [alias]' }, {
        option: '--type [type]',
        autocomplete: ['PowerPoint', 'Word', 'Excel', 'Other']
    });
}, _PlannerTaskReferenceAddCommand_initValidators = function _PlannerTaskReferenceAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type && ['powerpoint', 'word', 'excel', 'other'].indexOf(args.options.type.toLocaleLowerCase()) === -1) {
            return `${args.options.type} is not a valid type value. Allowed values PowerPoint|Word|Excel|Other`;
        }
        return true;
    });
}, _PlannerTaskReferenceAddCommand_initTypes = function _PlannerTaskReferenceAddCommand_initTypes() {
    this.types.string.push('taskId', 'url', 'alias', 'type');
};
export default new PlannerTaskReferenceAddCommand();
//# sourceMappingURL=task-reference-add.js.map