var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerPlanSetCommand_instances, _PlannerPlanSetCommand_initTelemetry, _PlannerPlanSetCommand_initOptions, _PlannerPlanSetCommand_initValidators, _PlannerPlanSetCommand_initOptionSets, _PlannerPlanSetCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerPlanSetCommand extends GraphCommand {
    get name() {
        return commands.PLAN_SET;
    }
    get description() {
        return 'Updates a Microsoft Planner plan';
    }
    constructor() {
        super();
        _PlannerPlanSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerPlanSetCommand_instances, "m", _PlannerPlanSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerPlanSetCommand_instances, "m", _PlannerPlanSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerPlanSetCommand_instances, "m", _PlannerPlanSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerPlanSetCommand_instances, "m", _PlannerPlanSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerPlanSetCommand_instances, "m", _PlannerPlanSetCommand_initTypes).call(this);
    }
    allowUnknownOptions() {
        return true;
    }
    async getGroupId(args) {
        const { ownerGroupId, ownerGroupName } = args.options;
        if (ownerGroupId) {
            return ownerGroupId;
        }
        const id = await entraGroup.getGroupIdByDisplayName(ownerGroupName);
        return id;
    }
    async getPlanId(args) {
        const { id, title } = args.options;
        if (id) {
            return id;
        }
        if (args.options.rosterId) {
            const id = await planner.getPlanIdByRosterId(args.options.rosterId);
            return id;
        }
        else {
            const groupId = await this.getGroupId(args);
            const id = await planner.getPlanIdByTitle(title, groupId);
            return id;
        }
    }
    async getUserIds(options) {
        if (options.shareWithUserIds) {
            return options.shareWithUserIds.split(',');
        }
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
            throw `Cannot proceed with planner plan creation. The following users provided are invalid: ${invalidUsers.join(',')}`;
        }
        return userIds;
    }
    async generateSharedWith(options) {
        const sharedWith = {};
        const userIds = await this.getUserIds(options);
        userIds.map(x => sharedWith[x] = true);
        return sharedWith;
    }
    async getPlanEtag(planId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/plans/${planId}`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response['@odata.etag'];
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
    async getPlanDetails(plan) {
        const requestOptionsTaskDetails = {
            url: `${this.resource}/v1.0/planner/plans/${plan.id}/details`,
            headers: {
                'accept': 'application/json;odata.metadata=none',
                'Prefer': 'return=representation'
            },
            responseType: 'json'
        };
        const planDetails = await request.get(requestOptionsTaskDetails);
        return { ...plan, ...planDetails };
    }
    async updatePlanDetails(options, planId) {
        const plan = await planner.getPlanById(planId);
        const categories = {};
        let categoriesCount = 0;
        Object.keys(options).forEach(key => {
            if (key.indexOf('category') !== -1) {
                categories[key] = options[key];
                categoriesCount++;
            }
        });
        if (!options.shareWithUserIds && !options.shareWithUserNames && categoriesCount === 0) {
            return this.getPlanDetails(plan);
        }
        const requestBody = {};
        if (options.shareWithUserIds || options.shareWithUserNames) {
            const sharedWith = await this.generateSharedWith(options);
            requestBody['sharedWith'] = sharedWith;
        }
        if (categoriesCount > 0) {
            requestBody['categoryDescriptions'] = categories;
        }
        const etag = await this.getPlanDetailsEtag(planId);
        const requestOptionsPlanDetails = {
            url: `${this.resource}/v1.0/planner/plans/${planId}/details`,
            headers: {
                'accept': 'application/json;odata.metadata=none',
                'If-Match': etag,
                'Prefer': 'return=representation'
            },
            responseType: 'json',
            data: requestBody
        };
        const planDetails = await request.patch(requestOptionsPlanDetails);
        return { ...plan, ...planDetails };
    }
    async commandAction(logger, args) {
        try {
            const planId = await this.getPlanId(args);
            if (args.options.newTitle) {
                const etag = await this.getPlanEtag(planId);
                const requestOptions = {
                    url: `${this.resource}/v1.0/planner/plans/${planId}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none',
                        'If-Match': etag,
                        'Prefer': 'return=representation'
                    },
                    responseType: 'json',
                    data: {
                        "title": args.options.newTitle
                    }
                };
                await request.patch(requestOptions);
            }
            const result = await this.updatePlanDetails(args.options, planId);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PlannerPlanSetCommand_instances = new WeakSet(), _PlannerPlanSetCommand_initTelemetry = function _PlannerPlanSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            newTitle: typeof args.options.newTitle !== 'undefined',
            shareWithUserIds: typeof args.options.shareWithUserIds !== 'undefined',
            shareWithUserNames: typeof args.options.shareWithUserNames !== 'undefined'
        });
    });
}, _PlannerPlanSetCommand_initOptions = function _PlannerPlanSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '--ownerGroupId [ownerGroupId]'
    }, {
        option: '--ownerGroupName [ownerGroupName]'
    }, {
        option: '--rosterId [rosterId]'
    }, {
        option: '--newTitle [newTitle]'
    }, {
        option: '--shareWithUserIds [shareWithUserIds]'
    }, {
        option: '--shareWithUserNames [shareWithUserNames]'
    });
}, _PlannerPlanSetCommand_initValidators = function _PlannerPlanSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.title) {
            if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
                return `${args.options.ownerGroupId} is not a valid GUID`;
            }
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
        const allowedCategories = [
            'category1',
            'category2',
            'category3',
            'category4',
            'category5',
            'category6',
            'category7',
            'category8',
            'category9',
            'category10',
            'category11',
            'category12',
            'category13',
            'category14',
            'category15',
            'category16',
            'category17',
            'category18',
            'category19',
            'category20',
            'category21',
            'category22',
            'category23',
            'category24',
            'category25'
        ];
        let invalidCategoryOptions = false;
        Object.keys(args.options).forEach(key => {
            if (key.indexOf('category') !== -1 && allowedCategories.indexOf(key) === -1) {
                invalidCategoryOptions = true;
            }
        });
        if (invalidCategoryOptions) {
            return 'Specify category values between category1 to category25';
        }
        return true;
    });
}, _PlannerPlanSetCommand_initOptionSets = function _PlannerPlanSetCommand_initOptionSets() {
    this.optionSets.push({
        options: ['id', 'title', 'rosterId']
    }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => {
            return args.options.title !== undefined;
        }
    });
}, _PlannerPlanSetCommand_initTypes = function _PlannerPlanSetCommand_initTypes() {
    this.types.string.push('id', 'title', 'ownerGroupId', 'ownerGroupName', 'rosterId', 'newTitle', 'shareWithUserIds', 'shareWithUserNames');
};
export default new PlannerPlanSetCommand();
//# sourceMappingURL=plan-set.js.map