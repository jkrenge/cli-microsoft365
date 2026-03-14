var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskAddCommand_instances, _PlannerTaskAddCommand_initTelemetry, _PlannerTaskAddCommand_initOptions, _PlannerTaskAddCommand_initValidators, _PlannerTaskAddCommand_initOptionSets, _PlannerTaskAddCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { taskPriority } from '../../taskPriority.js';
class PlannerTaskAddCommand extends GraphCommand {
    get name() {
        return commands.TASK_ADD;
    }
    get description() {
        return 'Adds a new Microsoft Planner Task';
    }
    constructor() {
        super();
        _PlannerTaskAddCommand_instances.add(this);
        this.allowedAppliedCategories = ['category1', 'category2', 'category3', 'category4', 'category5', 'category6'];
        this.allowedPreviewTypes = ['automatic', 'nopreview', 'checklist', 'description', 'reference'];
        __classPrivateFieldGet(this, _PlannerTaskAddCommand_instances, "m", _PlannerTaskAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTaskAddCommand_instances, "m", _PlannerTaskAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskAddCommand_instances, "m", _PlannerTaskAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerTaskAddCommand_instances, "m", _PlannerTaskAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerTaskAddCommand_instances, "m", _PlannerTaskAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            this.planId = await this.getPlanId(args);
            this.bucketId = await this.getBucketId(args, this.planId);
            const assignments = await this.generateUserAssignments(args);
            const appliedCategories = this.generateAppliedCategories(args.options);
            const requestOptions = {
                url: `${this.resource}/v1.0/planner/tasks`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    planId: this.planId,
                    bucketId: this.bucketId,
                    title: args.options.title,
                    startDateTime: args.options.startDateTime,
                    dueDateTime: args.options.dueDateTime,
                    percentComplete: args.options.percentComplete,
                    assignments: assignments,
                    orderHint: args.options.orderHint,
                    assigneePriority: args.options.assigneePriority,
                    appliedCategories: appliedCategories,
                    priority: taskPriority.getPriorityValue(args.options.priority)
                }
            };
            const newTask = await request.post(requestOptions);
            const result = await this.updateTaskDetails(args.options, newTask);
            if (result.description) {
                result.hasDescription = true;
            }
            await logger.log(result);
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
    generateAppliedCategories(options) {
        if (!options.appliedCategories) {
            return {};
        }
        const categories = {};
        options.appliedCategories.toLocaleLowerCase().split(',').forEach(x => categories[x] = true);
        return categories;
    }
    async updateTaskDetails(options, newTask) {
        const taskId = newTask.id;
        if (!options.description && !options.previewType) {
            return newTask;
        }
        const etag = await this.getTaskDetailsEtag(taskId);
        const requestOptionsTaskDetails = {
            url: `${this.resource}/v1.0/planner/tasks/${taskId}/details`,
            headers: {
                'accept': 'application/json;odata.metadata=none',
                'If-Match': etag,
                'Prefer': 'return=representation'
            },
            responseType: 'json',
            data: {
                description: options.description,
                previewType: options.previewType
            }
        };
        const taskDetails = await request.patch(requestOptionsTaskDetails);
        return { ...newTask, ...taskDetails };
    }
    async generateUserAssignments(args) {
        const assignments = {};
        if (!args.options.assignedToUserIds && !args.options.assignedToUserNames) {
            return assignments;
        }
        const userIds = await this.getUserIds(args.options);
        userIds.map(x => assignments[x] = {
            '@odata.type': '#microsoft.graph.plannerAssignment',
            orderHint: ' !'
        });
        return assignments;
    }
    async getBucketId(args, planId) {
        if (args.options.bucketId) {
            return args.options.bucketId;
        }
        return planner.getBucketIdByTitle(args.options.bucketName, planId);
    }
    async getPlanId(args) {
        if (args.options.planId) {
            return args.options.planId;
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
            return args.options.ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName);
    }
    async getUserIds(options) {
        if (options.assignedToUserIds) {
            return options.assignedToUserIds.split(',');
        }
        // Hitting this section means assignedToUserNames won't be undefined
        const userNames = options.assignedToUserNames;
        const userArr = userNames.split(',').map(o => o.trim());
        const promises = userArr.map(user => {
            const requestOptions = {
                url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
                headers: {
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            return request.get(requestOptions);
        });
        const usersRes = await Promise.all(promises);
        let userUpns = [];
        userUpns = usersRes.map(res => res.value[0]?.userPrincipalName);
        const userIds = usersRes.map(res => res.value[0]?.id);
        // Find the members where no graph response was found
        const invalidUsers = userArr.filter(user => !userUpns.some((upn) => upn?.toLowerCase() === user.toLowerCase()));
        if (invalidUsers && invalidUsers.length > 0) {
            throw `Cannot proceed with planner task creation. The following users provided are invalid : ${invalidUsers.join(',')}`;
        }
        return userIds;
    }
}
_PlannerTaskAddCommand_instances = new WeakSet(), _PlannerTaskAddCommand_initTelemetry = function _PlannerTaskAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            planId: typeof args.options.planId !== 'undefined',
            planTitle: typeof args.options.planTitle !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
            bucketId: typeof args.options.bucketId !== 'undefined',
            bucketName: typeof args.options.bucketName !== 'undefined',
            startDateTime: typeof args.options.startDateTime !== 'undefined',
            dueDateTime: typeof args.options.dueDateTime !== 'undefined',
            percentComplete: typeof args.options.percentComplete !== 'undefined',
            assignedToUserIds: typeof args.options.assignedToUserIds !== 'undefined',
            assignedToUserNames: typeof args.options.assignedToUserNames !== 'undefined',
            assigneePriority: typeof args.options.assigneePriority !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            appliedCategories: typeof args.options.appliedCategories !== 'undefined',
            previewType: typeof args.options.previewType !== 'undefined',
            orderHint: typeof args.options.orderHint !== 'undefined',
            priority: typeof args.options.priority !== 'undefined'
        });
    });
}, _PlannerTaskAddCommand_initOptions = function _PlannerTaskAddCommand_initOptions() {
    this.options.unshift({ option: '-t, --title <title>' }, { option: '--planId [planId]' }, { option: '--planTitle [planTitle]' }, { option: '--rosterId [rosterId]' }, { option: '--ownerGroupId [ownerGroupId]' }, { option: '--ownerGroupName [ownerGroupName]' }, { option: '--bucketId [bucketId]' }, { option: '--bucketName [bucketName]' }, { option: '--startDateTime [startDateTime]' }, { option: '--dueDateTime [dueDateTime]' }, { option: '--percentComplete [percentComplete]' }, { option: '--assignedToUserIds [assignedToUserIds]' }, { option: '--assignedToUserNames [assignedToUserNames]' }, { option: '--assigneePriority [assigneePriority]' }, { option: '--description [description]' }, {
        option: '--appliedCategories [appliedCategories]',
        autocomplete: this.allowedAppliedCategories
    }, {
        option: '--previewType [previewType]',
        autocomplete: this.allowedPreviewTypes
    }, { option: '--orderHint [orderHint]' }, { option: '--priority [priority]', autocomplete: taskPriority.priorityValues });
}, _PlannerTaskAddCommand_initValidators = function _PlannerTaskAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
        }
        if (args.options.startDateTime && !validation.isValidISODateTime(args.options.startDateTime)) {
            return 'The startDateTime is not a valid ISO date string';
        }
        if (args.options.dueDateTime && !validation.isValidISODateTime(args.options.dueDateTime)) {
            return 'The dueDateTime is not a valid ISO date string';
        }
        if (args.options.percentComplete && isNaN(args.options.percentComplete)) {
            return `percentComplete is not a number`;
        }
        if (args.options.percentComplete && (args.options.percentComplete < 0 || args.options.percentComplete > 100)) {
            return `percentComplete should be between 0 and 100`;
        }
        if (args.options.assignedToUserIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.assignedToUserIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'assignedToUserIds': ${isValidGUIDArrayResult}.`;
            }
        }
        if (args.options.assignedToUserIds && args.options.assignedToUserNames) {
            return 'Specify either assignedToUserIds or assignedToUserNames but not both';
        }
        if (args.options.assignedToUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.assignedToUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'assignedToUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.appliedCategories && args.options.appliedCategories.split(',').filter(category => this.allowedAppliedCategories.indexOf(category.toLocaleLowerCase()) < 0).length !== 0) {
            return `The appliedCategories contains invalid value. Specify either ${this.allowedAppliedCategories.join(', ')} as properties`;
        }
        if (args.options.previewType && this.allowedPreviewTypes.indexOf(args.options.previewType.toLocaleLowerCase()) === -1) {
            return `${args.options.previewType} is not a valid preview type value. Allowed values are ${this.allowedPreviewTypes.join(', ')}`;
        }
        if (args.options.priority !== undefined) {
            if (typeof args.options.priority === "number") {
                if (isNaN(args.options.priority) || args.options.priority < 0 || args.options.priority > 10 || !Number.isInteger(args.options.priority)) {
                    return 'priority should be an integer between 0 and 10.';
                }
            }
            else if (taskPriority.priorityValues.map(l => l.toLowerCase()).indexOf(args.options.priority.toString().toLowerCase()) === -1) {
                return `${args.options.priority} is not a valid priority value. Allowed values are ${taskPriority.priorityValues.join('|')}.`;
            }
        }
        return true;
    });
}, _PlannerTaskAddCommand_initOptionSets = function _PlannerTaskAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['planId', 'planTitle', 'rosterId'] }, { options: ['bucketId', 'bucketName'] }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => {
            return args.options.planTitle !== undefined;
        }
    });
}, _PlannerTaskAddCommand_initTypes = function _PlannerTaskAddCommand_initTypes() {
    this.types.string.push('title', 'planId', 'planTitle', 'rosterId', 'ownerGroupId', 'ownerGroupName', 'bucketId', 'bucketName', 'startDateTime', 'dueDateTime', 'assignedToUserIds', 'assignedToUserNames', 'appliedCategories', 'previewType', 'description', 'assigneePriority', 'orderHint');
};
export default new PlannerTaskAddCommand();
//# sourceMappingURL=task-add.js.map