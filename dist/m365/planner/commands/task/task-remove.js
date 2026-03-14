var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskRemoveCommand_instances, _PlannerTaskRemoveCommand_initTelemetry, _PlannerTaskRemoveCommand_initOptions, _PlannerTaskRemoveCommand_initValidators, _PlannerTaskRemoveCommand_initOptionSets, _PlannerTaskRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { formatting } from '../../../../utils/formatting.js';
class PlannerTaskRemoveCommand extends GraphCommand {
    get name() {
        return commands.TASK_REMOVE;
    }
    get description() {
        return 'Removes the Microsoft Planner task from a plan';
    }
    constructor() {
        super();
        _PlannerTaskRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTaskRemoveCommand_instances, "m", _PlannerTaskRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTaskRemoveCommand_instances, "m", _PlannerTaskRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskRemoveCommand_instances, "m", _PlannerTaskRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerTaskRemoveCommand_instances, "m", _PlannerTaskRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerTaskRemoveCommand_instances, "m", _PlannerTaskRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeTask = async () => {
            try {
                const task = await this.getTask(args.options);
                if (this.verbose) {
                    await logger.logToStderr(`Removing task '${task.title}' ...`);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/planner/tasks/${task.id}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none',
                        'if-match': task['@odata.etag']
                    },
                    responseType: 'json'
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeTask();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the task ${args.options.id || args.options.title}?` });
            if (result) {
                await removeTask();
            }
        }
    }
    async getTask(options) {
        const { id, title } = options;
        if (id) {
            const requestOptions = {
                url: `${this.resource}/v1.0/planner/tasks/${id}`,
                headers: {
                    accept: 'application/json'
                },
                responseType: 'json'
            };
            return await request.get(requestOptions);
        }
        const bucketId = await this.getBucketId(options);
        // $filter is not working on the buckets/{bucketId}/tasks endpoint, hence it is not being used.
        const tasks = await odata.getAllItems(`${this.resource}/v1.0/planner/buckets/${bucketId}/tasks?$select=title,id`, 'minimal');
        const filteredTasks = tasks.filter(b => title.toLocaleLowerCase() === b.title.toLocaleLowerCase());
        if (filteredTasks.length === 0) {
            throw `The specified task ${title} does not exist`;
        }
        if (filteredTasks.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', filteredTasks);
            return await cli.handleMultipleResultsFound(`Multiple tasks with title '${title}' found.`, resultAsKeyValuePair);
        }
        return filteredTasks[0];
    }
    async getBucketId(options) {
        const { bucketId, bucketName } = options;
        if (bucketId) {
            return bucketId;
        }
        const planId = await this.getPlanId(options);
        return planner.getBucketIdByTitle(bucketName, planId);
    }
    async getPlanId(options) {
        const { planId, planTitle, rosterId } = options;
        if (planId) {
            return planId;
        }
        if (options.rosterId) {
            return planner.getPlanIdByRosterId(rosterId);
        }
        else {
            const groupId = await this.getGroupId(options);
            return planner.getPlanIdByTitle(planTitle, groupId);
        }
    }
    async getGroupId(options) {
        const { ownerGroupId, ownerGroupName } = options;
        if (ownerGroupId) {
            return ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(ownerGroupName);
    }
}
_PlannerTaskRemoveCommand_instances = new WeakSet(), _PlannerTaskRemoveCommand_initTelemetry = function _PlannerTaskRemoveCommand_initTelemetry() {
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
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _PlannerTaskRemoveCommand_initOptions = function _PlannerTaskRemoveCommand_initOptions() {
    this.options.unshift({ option: '-i, --id [id]' }, { option: '-t, --title [title]' }, { option: '--bucketId [bucketId]' }, { option: '--bucketName [bucketName]' }, { option: '--planId [planId]' }, { option: '--planTitle [planTitle]' }, { option: '--rosterId [rosterId]' }, { option: '--ownerGroupId [ownerGroupId]' }, { option: '--ownerGroupName [ownerGroupName]' }, { option: '-f, --force' });
}, _PlannerTaskRemoveCommand_initValidators = function _PlannerTaskRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id) {
            if (args.options.bucketId || args.options.bucketName ||
                args.options.planId || args.options.planTitle || args.options.rosterId ||
                args.options.ownerGroupId || args.options.ownerGroupName) {
                return 'Don\'t specify bucketId,bucketName, planId, planTitle, rosterId, ownerGroupId, or ownerGroupName when using id';
            }
        }
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
        }
        return true;
    });
}, _PlannerTaskRemoveCommand_initOptionSets = function _PlannerTaskRemoveCommand_initOptionSets() {
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
}, _PlannerTaskRemoveCommand_initTypes = function _PlannerTaskRemoveCommand_initTypes() {
    this.types.string.push('id', 'title', 'planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'bucketId', 'bucketName', 'rosterId');
};
export default new PlannerTaskRemoveCommand();
//# sourceMappingURL=task-remove.js.map