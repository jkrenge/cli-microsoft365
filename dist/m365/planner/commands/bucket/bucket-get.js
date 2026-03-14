var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerBucketGetCommand_instances, _PlannerBucketGetCommand_initTelemetry, _PlannerBucketGetCommand_initOptions, _PlannerBucketGetCommand_initValidators, _PlannerBucketGetCommand_initOptionSets, _PlannerBucketGetCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerBucketGetCommand extends GraphCommand {
    get name() {
        return commands.BUCKET_GET;
    }
    get description() {
        return 'Gets the Microsoft Planner bucket in a plan';
    }
    constructor() {
        super();
        _PlannerBucketGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerBucketGetCommand_instances, "m", _PlannerBucketGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerBucketGetCommand_instances, "m", _PlannerBucketGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerBucketGetCommand_instances, "m", _PlannerBucketGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerBucketGetCommand_instances, "m", _PlannerBucketGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerBucketGetCommand_instances, "m", _PlannerBucketGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const bucket = await this.getBucket(args);
            await logger.log(bucket);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getBucket(args) {
        const { id, name } = args.options;
        if (id) {
            return this.getBucketById(id);
        }
        const planId = await this.getPlanId(args);
        return planner.getBucketByTitle(name, planId);
    }
    async getPlanId(args) {
        const { planId, planTitle, rosterId } = args.options;
        if (planId) {
            return planId;
        }
        if (planTitle) {
            const groupId = await this.getGroupId(args);
            return planner.getPlanIdByTitle(planTitle, groupId);
        }
        return planner.getPlanIdByRosterId(rosterId);
    }
    async getBucketById(id) {
        const requestOptions = {
            url: `${this.resource}/v1.0/planner/buckets/${id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    async getGroupId(args) {
        const { ownerGroupId, ownerGroupName } = args.options;
        if (ownerGroupId) {
            return ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(ownerGroupName);
    }
}
_PlannerBucketGetCommand_instances = new WeakSet(), _PlannerBucketGetCommand_initTelemetry = function _PlannerBucketGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            planId: typeof args.options.planId !== 'undefined',
            planTitle: typeof args.options.planTitle !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined'
        });
    });
}, _PlannerBucketGetCommand_initOptions = function _PlannerBucketGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--planId [planId]'
    }, {
        option: "--planTitle [planTitle]"
    }, {
        option: '--rosterId [rosterId]'
    }, {
        option: '--ownerGroupId [ownerGroupId]'
    }, {
        option: '--ownerGroupName [ownerGroupName]'
    });
}, _PlannerBucketGetCommand_initValidators = function _PlannerBucketGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id) {
            if (args.options.planId || args.options.planTitle || args.options.rosterId || args.options.ownerGroupId || args.options.ownerGroupName) {
                return 'Don\'t specify planId, planTitle, rosterId, ownerGroupId or ownerGroupName when using id';
            }
        }
        if (args.options.name) {
            if (args.options.planTitle) {
                if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
                    return `${args.options.ownerGroupId} is not a valid GUID`;
                }
            }
            if (args.options.planId) {
                if (args.options.ownerGroupId || args.options.ownerGroupName) {
                    return 'Don\'t specify ownerGroupId or ownerGroupName when using planId';
                }
            }
            if (args.options.rosterId) {
                if (args.options.ownerGroupId || args.options.ownerGroupName) {
                    return 'Don\'t specify ownerGroupId or ownerGroupName when using rosterId';
                }
            }
        }
        return true;
    });
}, _PlannerBucketGetCommand_initOptionSets = function _PlannerBucketGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] }, {
        options: ['planId', 'planTitle', 'rosterId'],
        runsWhen: (args) => args.options.name !== undefined
    }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => args.options.planTitle !== undefined
    });
}, _PlannerBucketGetCommand_initTypes = function _PlannerBucketGetCommand_initTypes() {
    this.types.string.push('id', 'name', 'planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'rosterId ');
};
export default new PlannerBucketGetCommand();
//# sourceMappingURL=bucket-get.js.map