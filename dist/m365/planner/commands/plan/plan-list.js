var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerPlanListCommand_instances, _PlannerPlanListCommand_initTelemetry, _PlannerPlanListCommand_initOptions, _PlannerPlanListCommand_initValidators, _PlannerPlanListCommand_initOptionSets, _PlannerPlanListCommand_initTypes;
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerPlanListCommand extends GraphCommand {
    get name() {
        return commands.PLAN_LIST;
    }
    get description() {
        return 'Returns a list of plans associated with a specified group or roster';
    }
    constructor() {
        super();
        _PlannerPlanListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerPlanListCommand_instances, "m", _PlannerPlanListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerPlanListCommand_instances, "m", _PlannerPlanListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerPlanListCommand_instances, "m", _PlannerPlanListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerPlanListCommand_instances, "m", _PlannerPlanListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerPlanListCommand_instances, "m", _PlannerPlanListCommand_initTypes).call(this);
    }
    defaultProperties() {
        return ['id', 'title', 'createdDateTime', 'owner'];
    }
    async commandAction(logger, args) {
        try {
            let plannerPlans = [];
            if (args.options.ownerGroupId || args.options.ownerGroupName) {
                const groupId = await this.getGroupId(args);
                plannerPlans = await planner.getPlansByGroupId(groupId);
            }
            else {
                const plan = await planner.getPlanByRosterId(args.options.rosterId);
                plannerPlans.push(plan);
            }
            await logger.log(plannerPlans);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getGroupId(args) {
        if (args.options.ownerGroupId) {
            return args.options.ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName);
    }
}
_PlannerPlanListCommand_instances = new WeakSet(), _PlannerPlanListCommand_initTelemetry = function _PlannerPlanListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined'
        });
    });
}, _PlannerPlanListCommand_initOptions = function _PlannerPlanListCommand_initOptions() {
    this.options.unshift({
        option: "--ownerGroupId [ownerGroupId]"
    }, {
        option: "--ownerGroupName [ownerGroupName]"
    }, {
        option: "--rosterId [rosterId]"
    });
}, _PlannerPlanListCommand_initValidators = function _PlannerPlanListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
        }
        return true;
    });
}, _PlannerPlanListCommand_initOptionSets = function _PlannerPlanListCommand_initOptionSets() {
    this.optionSets.push({ options: ['ownerGroupId', 'ownerGroupName', 'rosterId'] });
}, _PlannerPlanListCommand_initTypes = function _PlannerPlanListCommand_initTypes() {
    this.types.string.push('ownerGroupId', 'ownerGroupName', 'rosterId');
};
export default new PlannerPlanListCommand();
//# sourceMappingURL=plan-list.js.map