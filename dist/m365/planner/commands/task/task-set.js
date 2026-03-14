var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerTaskSetCommand_instances, _PlannerTaskSetCommand_initTelemetry, _PlannerTaskSetCommand_initOptions, _PlannerTaskSetCommand_initValidators, _PlannerTaskSetCommand_initOptionSets, _PlannerTaskSetCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { taskPriority } from '../../taskPriority.js';
class PlannerTaskSetCommand extends GraphCommand {
    get name() {
        return commands.TASK_SET;
    }
    get description() {
        return 'Updates a Microsoft Planner Task';
    }
    constructor() {
        super();
        _PlannerTaskSetCommand_instances.add(this);
        this.allowedAppliedCategories = ['category1', 'category2', 'category3', 'category4', 'category5', 'category6'];
        __classPrivateFieldGet(this, _PlannerTaskSetCommand_instances, "m", _PlannerTaskSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerTaskSetCommand_instances, "m", _PlannerTaskSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerTaskSetCommand_instances, "m", _PlannerTaskSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerTaskSetCommand_instances, "m", _PlannerTaskSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerTaskSetCommand_instances, "m", _PlannerTaskSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            this.bucketId = await this.getBucketId(args.options);
            this.assignments = await this.generateUserAssignments(args.options);
            const etag = await this.getTaskEtag(args.options.id);
            const appliedCategories = this.generateAppliedCategories(args.options);
            const data = this.mapRequestBody(args.options, appliedCategories);
            const requestOptions = {
                url: `${this.resource}/v1.0/planner/tasks/${args.options.id}`,
                headers: {
                    'accept': 'application/json;odata.metadata=none',
                    'If-Match': etag,
                    'Prefer': 'return=representation'
                },
                responseType: 'json',
                data: data
            };
            const newTask = await request.patch(requestOptions);
            const result = await this.updateTaskDetails(args.options, newTask);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async updateTaskDetails(options, newTask) {
        if (!options.description) {
            return newTask;
        }
        const taskId = newTask.id;
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
                description: options.description
            }
        };
        const taskDetails = await request.patch(requestOptionsTaskDetails);
        return { ...newTask, ...taskDetails };
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
    async getTaskEtag(taskId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/tasks/${formatting.encodeQueryParameter(taskId)}`,
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
    async generateUserAssignments(options) {
        const assignments = {};
        if (!options.assignedToUserIds && !options.assignedToUserNames) {
            return assignments;
        }
        const userIds = await this.getUserIds(options);
        userIds.forEach(x => assignments[x] = {
            '@odata.type': '#microsoft.graph.plannerAssignment',
            orderHint: ' !'
        });
        return assignments;
    }
    async getUserIds(options) {
        if (options.assignedToUserIds) {
            return options.assignedToUserIds.split(',').map(o => o.trim());
        }
        // Hitting this section means assignedToUserNames won't be undefined
        const userNames = options.assignedToUserNames;
        const userArr = userNames.split(',').map(o => o.trim());
        const promises = userArr.map(user => {
            const requestOptions = {
                url: `${this.resource}/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`,
                headers: {
                    'accept ': 'application/json;odata.metadata=none'
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
            throw `Cannot proceed with planner task update. The following users provided are invalid : ${invalidUsers.join(',')}`;
        }
        return userIds;
    }
    async getBucketId(options) {
        if (options.bucketId) {
            return options.bucketId;
        }
        if (!options.bucketName) {
            return undefined;
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
    mapRequestBody(options, appliedCategories) {
        const requestBody = {};
        if (options.title) {
            requestBody.title = options.title;
        }
        if (this.bucketId) {
            requestBody.bucketId = this.bucketId;
        }
        if (options.startDateTime) {
            requestBody.startDateTime = options.startDateTime;
        }
        if (options.dueDateTime) {
            requestBody.dueDateTime = options.dueDateTime;
        }
        if (options.percentComplete) {
            requestBody.percentComplete = options.percentComplete;
        }
        if (this.assignments && Object.keys(this.assignments).length > 0) {
            requestBody.assignments = this.assignments;
        }
        if (options.assigneePriority) {
            requestBody.assigneePriority = options.assigneePriority;
        }
        if (appliedCategories && Object.keys(appliedCategories).length > 0) {
            requestBody.appliedCategories = appliedCategories;
        }
        if (options.orderHint) {
            requestBody.orderHint = options.orderHint;
        }
        if (options.priority !== undefined) {
            requestBody.priority = taskPriority.getPriorityValue(options.priority);
        }
        return requestBody;
    }
}
_PlannerTaskSetCommand_instances = new WeakSet(), _PlannerTaskSetCommand_initTelemetry = function _PlannerTaskSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
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
            orderHint: typeof args.options.orderHint !== 'undefined',
            priority: typeof args.options.priority !== 'undefined'
        });
    });
}, _PlannerTaskSetCommand_initOptions = function _PlannerTaskSetCommand_initOptions() {
    this.options.unshift({ option: '-i, --id <id>' }, { option: '-t, --title [title]' }, { option: '--planId [planId]' }, { option: '--planTitle [planTitle]' }, { option: '--rosterId [rosterId]' }, { option: '--ownerGroupId [ownerGroupId]' }, { option: '--ownerGroupName [ownerGroupName]' }, { option: '--bucketId [bucketId]' }, { option: '--bucketName [bucketName]' }, { option: '--startDateTime [startDateTime]' }, { option: '--dueDateTime [dueDateTime]' }, { option: '--percentComplete [percentComplete]' }, { option: '--assignedToUserIds [assignedToUserIds]' }, { option: '--assignedToUserNames [assignedToUserNames]' }, { option: '--assigneePriority [assigneePriority]' }, { option: '--description [description]' }, { option: '--appliedCategories [appliedCategories]' }, { option: '--orderHint [orderHint]' }, { option: '--priority [priority]', autocomplete: taskPriority.priorityValues });
}, _PlannerTaskSetCommand_initValidators = function _PlannerTaskSetCommand_initValidators() {
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
        if (args.options.assignedToUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.assignedToUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'assignedToUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        if (args.options.appliedCategories && args.options.appliedCategories.split(',').filter(category => this.allowedAppliedCategories.indexOf(category.toLocaleLowerCase()) < 0).length !== 0) {
            return 'The appliedCategories contains invalid value. Specify either category1, category2, category3, category4, category5 and/or category6 as properties';
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
}, _PlannerTaskSetCommand_initOptionSets = function _PlannerTaskSetCommand_initOptionSets() {
    this.optionSets.push({
        options: ['bucketId', 'bucketName'],
        runsWhen: (args) => {
            return args.options.bucketId || args.options.bucketName;
        }
    }, {
        options: ['planId', 'planTitle', 'rosterId'],
        runsWhen: (args) => {
            return args.options.bucketName !== undefined;
        }
    }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => {
            return args.options.planTitle !== undefined;
        }
    }, {
        options: ['assignedToUserIds', 'assignedToUserNames'],
        runsWhen: (args) => {
            return args.options.assignedToUserIds || args.options.assignedToUserNames;
        }
    });
}, _PlannerTaskSetCommand_initTypes = function _PlannerTaskSetCommand_initTypes() {
    this.types.string.push('id', 'title', 'planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'bucketId', 'bucketName', 'description', 'assigneePriority', 'orderHint', 'rosterId', 'startDateTime', 'dueDateTime', 'assignedToUserIds', 'assignedToUserNames', 'assignments', 'appliedCategories');
};
export default new PlannerTaskSetCommand();
//# sourceMappingURL=task-set.js.map