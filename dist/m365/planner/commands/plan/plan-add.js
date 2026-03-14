var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerPlanAddCommand_instances, _PlannerPlanAddCommand_initTelemetry, _PlannerPlanAddCommand_initOptions, _PlannerPlanAddCommand_initValidators, _PlannerPlanAddCommand_initOptionSets, _PlannerPlanAddCommand_initTypes;
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerPlanAddCommand extends GraphCommand {
    get name() {
        return commands.PLAN_ADD;
    }
    get description() {
        return 'Adds a new Microsoft Planner plan';
    }
    constructor() {
        super();
        _PlannerPlanAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerPlanAddCommand_instances, "m", _PlannerPlanAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerPlanAddCommand_instances, "m", _PlannerPlanAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerPlanAddCommand_instances, "m", _PlannerPlanAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerPlanAddCommand_instances, "m", _PlannerPlanAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerPlanAddCommand_instances, "m", _PlannerPlanAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const data = {
                title: args.options.title
            };
            if (args.options.rosterId) {
                data.container = {
                    url: `https://graph.microsoft.com/v1.0/planner/rosters/${args.options.rosterId}`
                };
            }
            else {
                const groupId = await this.getGroupId(args);
                data.container = {
                    url: `https://graph.microsoft.com/v1.0/groups/${groupId}`
                };
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/planner/plans`,
                headers: {
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: data
            };
            const newPlan = await request.post(requestOptions);
            const result = await this.updatePlanDetails(args.options, newPlan);
            await logger.log(result);
        }
        catch (err) {
            if (args.options.rosterId && err.error?.error.message === "You do not have the required permissions to access this item, or the item may not exist.") {
                throw new CommandError("You can only add 1 plan to a Roster");
            }
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async updatePlanDetails(options, newPlan) {
        const planId = newPlan.id;
        if (!options.shareWithUserIds && !options.shareWithUserNames) {
            return newPlan;
        }
        const resArray = await Promise.all([this.generateSharedWith(options), this.getPlanDetailsEtag(planId)]);
        const sharedWith = resArray[0];
        const etag = resArray[1];
        const requestOptionsPlanDetails = {
            url: `${this.resource}/v1.0/planner/plans/${planId}/details`,
            headers: {
                'accept': 'application/json;odata.metadata=none',
                'If-Match': etag,
                'Prefer': 'return=representation'
            },
            responseType: 'json',
            data: {
                sharedWith: sharedWith
            }
        };
        const planDetails = await request.patch(requestOptionsPlanDetails);
        return { ...newPlan, ...planDetails };
    }
    async getPlanDetailsEtag(planId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/plans/${planId}/details`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response['@odata.etag'];
    }
    async generateSharedWith(options) {
        const sharedWith = {};
        const userIds = await this.getUserIds(options);
        userIds.map(x => sharedWith[x] = true);
        return sharedWith;
    }
    async getUserIds(options) {
        if (options.shareWithUserIds) {
            return options.shareWithUserIds.split(',');
        }
        // Hitting this section means assignedToUserNames won't be undefined
        const userNames = options.shareWithUserNames;
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
        const userUpns = usersRes.map(res => res.value[0]?.userPrincipalName);
        const userIds = usersRes.map(res => res.value[0]?.id);
        // Find the members where no graph response was found
        const invalidUsers = userArr.filter(user => !userUpns.some((upn) => upn?.toLowerCase() === user.toLowerCase()));
        if (invalidUsers && invalidUsers.length > 0) {
            throw `Cannot proceed with planner plan creation. The following users provided are invalid : ${invalidUsers.join(',')}`;
        }
        return userIds;
    }
    async getGroupId(args) {
        if (args.options.ownerGroupId) {
            return args.options.ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName);
    }
}
_PlannerPlanAddCommand_instances = new WeakSet(), _PlannerPlanAddCommand_initTelemetry = function _PlannerPlanAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
            shareWithUserIds: typeof args.options.shareWithUserIds !== 'undefined',
            shareWithUserNames: typeof args.options.shareWithUserNames !== 'undefined'
        });
    });
}, _PlannerPlanAddCommand_initOptions = function _PlannerPlanAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
    }, {
        option: "--ownerGroupId [ownerGroupId]"
    }, {
        option: "--ownerGroupName [ownerGroupName]"
    }, {
        option: "--rosterId [rosterId]"
    }, {
        option: '--shareWithUserIds [shareWithUserIds]'
    }, {
        option: '--shareWithUserNames [shareWithUserNames]'
    });
}, _PlannerPlanAddCommand_initValidators = function _PlannerPlanAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
        }
        if (args.options.shareWithUserIds && args.options.shareWithUserNames) {
            return 'Specify either shareWithUserIds or shareWithUserNames but not both';
        }
        if (args.options.shareWithUserIds) {
            const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.shareWithUserIds);
            if (isValidGUIDArrayResult !== true) {
                return `The following GUIDs are invalid for the option 'shareWithUserIds': ${isValidGUIDArrayResult}.`;
            }
        }
        if (args.options.shareWithUserNames) {
            const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.shareWithUserNames);
            if (isValidUPNArrayResult !== true) {
                return `The following user principal names are invalid for the option 'shareWithUserNames': ${isValidUPNArrayResult}.`;
            }
        }
        return true;
    });
}, _PlannerPlanAddCommand_initOptionSets = function _PlannerPlanAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['ownerGroupId', 'ownerGroupName', 'rosterId'] });
}, _PlannerPlanAddCommand_initTypes = function _PlannerPlanAddCommand_initTypes() {
    this.types.string.push('title', 'ownerGroupId', 'ownerGroupName', 'rosterId', 'shareWithUserIds', 'shareWithUserNames');
};
export default new PlannerPlanAddCommand();
//# sourceMappingURL=plan-add.js.map