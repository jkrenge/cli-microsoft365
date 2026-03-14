var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerBucketAddCommand_instances, _PlannerBucketAddCommand_initTelemetry, _PlannerBucketAddCommand_initOptions, _PlannerBucketAddCommand_initValidators, _PlannerBucketAddCommand_initOptionSets, _PlannerBucketAddCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerBucketAddCommand extends GraphCommand {
    get name() {
        return commands.BUCKET_ADD;
    }
    get description() {
        return 'Adds a new Microsoft Planner bucket';
    }
    constructor() {
        super();
        _PlannerBucketAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerBucketAddCommand_instances, "m", _PlannerBucketAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerBucketAddCommand_instances, "m", _PlannerBucketAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerBucketAddCommand_instances, "m", _PlannerBucketAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerBucketAddCommand_instances, "m", _PlannerBucketAddCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerBucketAddCommand_instances, "m", _PlannerBucketAddCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const planId = await this.getPlanId(args);
            const requestOptions = {
                url: `${this.resource}/v1.0/planner/buckets`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    name: args.options.name,
                    planId: planId,
                    orderHint: args.options.orderHint
                }
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
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
_PlannerBucketAddCommand_instances = new WeakSet(), _PlannerBucketAddCommand_initTelemetry = function _PlannerBucketAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            planId: typeof args.options.planId !== 'undefined',
            planTitle: typeof args.options.planTitle !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
            orderHint: typeof args.options.orderHint !== 'undefined'
        });
    });
}, _PlannerBucketAddCommand_initOptions = function _PlannerBucketAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: "--planId [planId]"
    }, {
        option: "--planTitle [planTitle]"
    }, {
        option: '--rosterId [rosterId]'
    }, {
        option: "--ownerGroupId [ownerGroupId]"
    }, {
        option: "--ownerGroupName [ownerGroupName]"
    }, {
        option: "--orderHint [orderHint]"
    });
}, _PlannerBucketAddCommand_initValidators = function _PlannerBucketAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
        }
        return true;
    });
}, _PlannerBucketAddCommand_initOptionSets = function _PlannerBucketAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['planId', 'planTitle', 'rosterId'] }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => args.options.planTitle !== undefined
    });
}, _PlannerBucketAddCommand_initTypes = function _PlannerBucketAddCommand_initTypes() {
    this.types.string.push('name', 'planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'orderHint', 'rosterId ');
};
export default new PlannerBucketAddCommand();
//# sourceMappingURL=bucket-add.js.map