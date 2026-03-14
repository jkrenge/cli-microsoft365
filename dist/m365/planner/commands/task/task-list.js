var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskListCommand_instances, _PlannerTaskListCommand_initTelemetry, _PlannerTaskListCommand_initOptions, _PlannerTaskListCommand_initValidators, _PlannerTaskListCommand_initOptionSets, _PlannerTaskListCommand_initTypes;
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerTaskListCommand extends GraphCommand {
    get name() {
        return commands.TASK_LIST;
    }
    get description() {
        return 'Lists planner tasks in a bucket, plan, or tasks for the currently logged in user';
    }
    defaultProperties() {
        return ['id', 'title', 'startDateTime', 'dueDateTime', 'completedDateTime'];
    }
    constructor() {
        super();
        _PlannerTaskListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerTaskListCommand_instances, "m", _PlannerTaskListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTaskListCommand_instances, "m", _PlannerTaskListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskListCommand_instances, "m", _PlannerTaskListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerTaskListCommand_instances, "m", _PlannerTaskListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerTaskListCommand_instances, "m", _PlannerTaskListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const bucketName = args.options.bucketName;
        let bucketId = args.options.bucketId;
        const planTitle = args.options.planTitle;
        let planId = args.options.planId;
        let taskItems;
        try {
            if (bucketId || bucketName) {
                bucketId = await this.getBucketId(args);
                taskItems = await odata.getAllItems(`${this.resource}/v1.0/planner/buckets/${bucketId}/tasks`);
                const betaTasks = await odata.getAllItems(`${this.resource}/beta/planner/buckets/${bucketId}/tasks`);
                await logger.log(this.mergeTaskPriority(taskItems, betaTasks));
            }
            else if (planId || planTitle) {
                planId = await this.getPlanId(args);
                taskItems = await odata.getAllItems(`${this.resource}/v1.0/planner/plans/${planId}/tasks`);
                const betaTasks = await odata.getAllItems(`${this.resource}/beta/planner/plans/${planId}/tasks`);
                await logger.log(this.mergeTaskPriority(taskItems, betaTasks));
            }
            else {
                taskItems = await odata.getAllItems(`${this.resource}/v1.0/me/planner/tasks`);
                const betaTasks = await odata.getAllItems(`${this.resource}/beta/me/planner/tasks`);
                await logger.log(this.mergeTaskPriority(taskItems, betaTasks));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getBucketId(args) {
        if (args.options.bucketId) {
            return formatting.encodeQueryParameter(args.options.bucketId);
        }
        const planId = await this.getPlanId(args);
        return planner.getBucketIdByTitle(args.options.bucketName, planId);
    }
    async getPlanId(args) {
        if (args.options.planId) {
            return formatting.encodeQueryParameter(args.options.planId);
        }
        if (args.options.rosterId) {
            return planner.getPlanIdByRosterId(args.options.rosterId);
        }
        else {
            const groupId = await this.getGroupId(args);
            return planner.getPlanIdByTitle(args.options.planTitle, groupId);
        }
    }
    async getGroupId(args) {
        if (args.options.ownerGroupId) {
            return formatting.encodeQueryParameter(args.options.ownerGroupId);
        }
        return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName);
    }
    mergeTaskPriority(taskItems, betaTaskItems) {
        const findBetaTask = (id) => betaTaskItems.find(task => task.id === id);
        taskItems.forEach(task => {
            const betaTaskItem = findBetaTask(task.id);
            if (betaTaskItem) {
                const { priority } = betaTaskItem;
                Object.assign(task, { priority });
            }
        });
        return taskItems;
    }
}
_PlannerTaskListCommand_instances = new WeakSet(), _PlannerTaskListCommand_initTelemetry = function _PlannerTaskListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            bucketId: typeof args.options.bucketId !== 'undefined',
            bucketName: typeof args.options.bucketName !== 'undefined',
            planId: typeof args.options.planId !== 'undefined',
            planTitle: typeof args.options.planTitle !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined'
        });
    });
}, _PlannerTaskListCommand_initOptions = function _PlannerTaskListCommand_initOptions() {
    this.options.unshift({
        option: '--bucketId [bucketId]'
    }, {
        option: '--bucketName [bucketName]'
    }, {
        option: '--planId [planId]'
    }, {
        option: '--planTitle [planTitle]'
    }, {
        option: '--rosterId [rosterId]'
    }, {
        option: '--ownerGroupId [ownerGroupId]'
    }, {
        option: '--ownerGroupName [ownerGroupName]'
    });
}, _PlannerTaskListCommand_initValidators = function _PlannerTaskListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
        }
        return true;
    });
}, _PlannerTaskListCommand_initOptionSets = function _PlannerTaskListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['bucketId', 'bucketName'],
        runsWhen: (args) => {
            return args.options.bucketId !== undefined || args.options.bucketName !== undefined;
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
}, _PlannerTaskListCommand_initTypes = function _PlannerTaskListCommand_initTypes() {
    this.types.string.push('planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'bucketId', 'bucketName', 'rosterId');
};
export default new PlannerTaskListCommand();
//# sourceMappingURL=task-list.js.map