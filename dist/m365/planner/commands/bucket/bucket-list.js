var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerBucketListCommand_instances, _PlannerBucketListCommand_initTelemetry, _PlannerBucketListCommand_initOptions, _PlannerBucketListCommand_initValidators, _PlannerBucketListCommand_initOptionSets, _PlannerBucketListCommand_initTypes;
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerBucketListCommand extends GraphCommand {
    get name() {
        return commands.BUCKET_LIST;
    }
    get description() {
        return 'Lists the Microsoft Planner buckets in a plan';
    }
    defaultProperties() {
        return ['id', 'name', 'planId', 'orderHint'];
    }
    constructor() {
        super();
        _PlannerBucketListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerBucketListCommand_instances, "m", _PlannerBucketListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerBucketListCommand_instances, "m", _PlannerBucketListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerBucketListCommand_instances, "m", _PlannerBucketListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerBucketListCommand_instances, "m", _PlannerBucketListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerBucketListCommand_instances, "m", _PlannerBucketListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const planId = await this.getPlanId(args);
            const buckets = await odata.getAllItems(`${this.resource}/v1.0/planner/plans/${planId}/buckets`);
            await logger.log(buckets);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getPlanId(args) {
        if (args.options.planId) {
            return args.options.planId;
        }
        if (args.options.planTitle) {
            const groupId = await this.getGroupId(args);
            return planner.getPlanIdByTitle(args.options.planTitle, groupId);
        }
        return planner.getPlanIdByRosterId(args.options.rosterId);
    }
    async getGroupId(args) {
        if (args.options.ownerGroupId) {
            return args.options.ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName);
    }
}
_PlannerBucketListCommand_instances = new WeakSet(), _PlannerBucketListCommand_initTelemetry = function _PlannerBucketListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            planId: typeof args.options.planId !== 'undefined',
            planTitle: typeof args.options.planTitle !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined'
        });
    });
}, _PlannerBucketListCommand_initOptions = function _PlannerBucketListCommand_initOptions() {
    this.options.unshift({
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
}, _PlannerBucketListCommand_initValidators = function _PlannerBucketListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
        }
        return true;
    });
}, _PlannerBucketListCommand_initOptionSets = function _PlannerBucketListCommand_initOptionSets() {
    this.optionSets.push({ options: ['planId', 'planTitle', 'rosterId'] }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => args.options.planTitle !== undefined
    });
}, _PlannerBucketListCommand_initTypes = function _PlannerBucketListCommand_initTypes() {
    this.types.string.push('planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'rosterId ');
};
export default new PlannerBucketListCommand();
//# sourceMappingURL=bucket-list.js.map