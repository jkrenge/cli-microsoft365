var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerBucketSetCommand_instances, _PlannerBucketSetCommand_initTelemetry, _PlannerBucketSetCommand_initOptions, _PlannerBucketSetCommand_initValidators, _PlannerBucketSetCommand_initOptionSets, _PlannerBucketSetCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerBucketSetCommand extends GraphCommand {
    get name() {
        return commands.BUCKET_SET;
    }
    get description() {
        return 'Updates a Microsoft Planner bucket';
    }
    constructor() {
        super();
        _PlannerBucketSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerBucketSetCommand_instances, "m", _PlannerBucketSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerBucketSetCommand_instances, "m", _PlannerBucketSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerBucketSetCommand_instances, "m", _PlannerBucketSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerBucketSetCommand_instances, "m", _PlannerBucketSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerBucketSetCommand_instances, "m", _PlannerBucketSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Updating bucket...`);
        }
        try {
            const bucket = await this.getBucket(args);
            const requestOptions = {
                url: `${this.resource}/v1.0/planner/buckets/${bucket.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'if-match': bucket['@odata.etag']
                },
                responseType: 'json',
                data: {
                    name: args.options.newName,
                    orderHint: args.options.orderHint
                }
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getBucket(args) {
        if (args.options.id) {
            const requestOptions = {
                url: `${this.resource}/v1.0/planner/buckets/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=minimal'
                },
                responseType: 'json'
            };
            return await request.get(requestOptions);
        }
        const planId = await this.getPlanId(args);
        return planner.getBucketByTitle(args.options.name, planId, 'minimal');
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
    async getGroupId(args) {
        const { ownerGroupId, ownerGroupName } = args.options;
        if (ownerGroupId) {
            return ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(ownerGroupName);
    }
}
_PlannerBucketSetCommand_instances = new WeakSet(), _PlannerBucketSetCommand_initTelemetry = function _PlannerBucketSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            planId: typeof args.options.planId !== 'undefined',
            planTitle: typeof args.options.planTitle !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
            newName: typeof args.options.newName !== 'undefined',
            orderHint: typeof args.options.orderHint !== 'undefined'
        });
    });
}, _PlannerBucketSetCommand_initOptions = function _PlannerBucketSetCommand_initOptions() {
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
    }, {
        option: '--newName [newName]'
    }, {
        option: '--orderHint [orderHint]'
    });
}, _PlannerBucketSetCommand_initValidators = function _PlannerBucketSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && (args.options.planId || args.options.planTitle || args.options.ownerGroupId || args.options.ownerGroupName || args.options.rosterId)) {
            return 'Don\'t specify planId, planTitle, ownerGroupId, ownerGroupName or rosterId when using id';
        }
        if (args.options.name) {
            if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
                return `${args.options.ownerGroupId} is not a valid GUID`;
            }
            if (args.options.planId && (args.options.ownerGroupId || args.options.ownerGroupName)) {
                return 'Don\'t specify ownerGroupId or ownerGroupName when using planId';
            }
            if (args.options.rosterId && (args.options.ownerGroupId || args.options.ownerGroupName)) {
                return 'Don\'t specify ownerGroupId or ownerGroupName when using rosterId';
            }
        }
        if (!args.options.newName && !args.options.orderHint) {
            return 'Specify either newName or orderHint';
        }
        return true;
    });
}, _PlannerBucketSetCommand_initOptionSets = function _PlannerBucketSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] }, {
        options: ['planId', 'planTitle', 'rosterId'],
        runsWhen: (args) => args.options.name !== undefined
    }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => args.options.planTitle !== undefined
    });
}, _PlannerBucketSetCommand_initTypes = function _PlannerBucketSetCommand_initTypes() {
    this.types.string.push('id', 'name', 'planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'orderHint', 'newName', 'rosterId ');
};
export default new PlannerBucketSetCommand();
//# sourceMappingURL=bucket-set.js.map