var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskGetCommand_instances, _PlannerTaskGetCommand_initTelemetry, _PlannerTaskGetCommand_initOptions, _PlannerTaskGetCommand_initValidators, _PlannerTaskGetCommand_initOptionSets, _PlannerTaskGetCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class PlannerTaskGetCommand extends GraphCommand {
    get name() {
        return commands.TASK_GET;
    }
    get description() {
        return 'Retrieve the specified planner task';
    }
    constructor() {
        super();
        _PlannerTaskGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTaskGetCommand_instances, "m", _PlannerTaskGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTaskGetCommand_instances, "m", _PlannerTaskGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskGetCommand_instances, "m", _PlannerTaskGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerTaskGetCommand_instances, "m", _PlannerTaskGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerTaskGetCommand_instances, "m", _PlannerTaskGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const taskId = await this.getTaskId(args.options);
            const task = await this.getTask(taskId);
            const res = await this.getTaskDetails(task);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTask(taskId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return await request.get(requestOptions);
    }
    async getTaskDetails(task) {
        const requestOptionsTaskDetails = {
            url: `${this.resource}/v1.0/planner/tasks/${task.id}/details`,
            headers: {
                'accept': 'application/json;odata.metadata=none',
                'Prefer': 'return=representation'
            },
            responseType: 'json'
        };
        const taskDetails = await request.get(requestOptionsTaskDetails);
        return { ...task, ...taskDetails };
    }
    async getTaskId(options) {
        if (options.id) {
            return options.id;
        }
        const bucketId = await this.getBucketId(options);
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/buckets/${bucketId}/tasks?$select=id,title`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const title = options.title;
        const tasks = response.value.filter(val => val.title?.toLocaleLowerCase() === title.toLocaleLowerCase());
        if (!tasks.length) {
            throw `The specified task ${options.title} does not exist`;
        }
        if (tasks.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', tasks);
            const result = await cli.handleMultipleResultsFound(`Multiple tasks with title '${options.title}' found.`, resultAsKeyValuePair);
            return result.id;
        }
        return tasks[0].id;
    }
    async getBucketId(options) {
        if (options.bucketId) {
            return options.bucketId;
        }
        const planId = await this.getPlanId(options);
        return planner.getBucketIdByTitle(options.bucketName, planId);
    }
    async getPlanId(options) {
        if (options.planId) {
            return options.planId;
        }
        if (options.rosterId) {
            return planner.getPlanIdByRosterId(options.rosterId);
        }
        else {
            const groupId = await this.getGroupId(options);
            return planner.getPlanIdByTitle(options.planTitle, groupId);
        }
    }
    async getGroupId(options) {
        if (options.ownerGroupId) {
            return options.ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(options.ownerGroupName);
    }
}
_PlannerTaskGetCommand_instances = new WeakSet(), _PlannerTaskGetCommand_initTelemetry = function _PlannerTaskGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            bucketId: typeof args.options.bucketId !== 'undefined',
            bucketName: typeof args.options.bucketName !== 'undefined',
            planId: typeof args.options.planId !== 'undefined',
            planTitle: typeof args.options.planTitle !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined'
        });
    });
}, _PlannerTaskGetCommand_initOptions = function _PlannerTaskGetCommand_initOptions() {
    this.options.unshift({ option: '-i, --id [id]' }, { option: '-t, --title [title]' }, { option: '--bucketId [bucketId]' }, { option: '--bucketName [bucketName]' }, { option: '--planId [planId]' }, { option: '--planTitle [planTitle]' }, { option: '--rosterId [rosterId]' }, { option: '--ownerGroupId [ownerGroupId]' }, { option: '--ownerGroupName [ownerGroupName]' });
}, _PlannerTaskGetCommand_initValidators = function _PlannerTaskGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id) {
            if (args.options.bucketId || args.options.bucketName ||
                args.options.planId || args.options.planTitle || args.options.rosterId ||
                args.options.ownerGroupId || args.options.ownerGroupName) {
                return 'Don\'t specify bucketId, bucketName, planId, planTitle, rosterId, ownerGroupId or ownerGroupName when using id';
            }
        }
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
        }
        return true;
    });
}, _PlannerTaskGetCommand_initOptionSets = function _PlannerTaskGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title'] }, {
        options: ['planId', 'planTitle', 'rosterId'],
        runsWhen: (args) => {
            return args.options.id === undefined;
        }
    }, {
        options: ['bucketId', 'bucketName'],
        runsWhen: (args) => {
            return args.options.title !== undefined;
        }
    }, {
        options: ['planId', 'planTitle'],
        runsWhen: (args) => {
            return args.options.bucketName !== undefined && args.options.rosterId === undefined;
        }
    }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => {
            return args.options.planTitle !== undefined;
        }
    });
}, _PlannerTaskGetCommand_initTypes = function _PlannerTaskGetCommand_initTypes() {
    this.types.string.push('id', 'title', 'planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'bucketId', 'bucketName', 'rosterId');
};
export default new PlannerTaskGetCommand();
//# sourceMappingURL=task-get.js.map