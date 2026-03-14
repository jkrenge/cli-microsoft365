var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerPlanGetCommand_instances, _PlannerPlanGetCommand_initTelemetry, _PlannerPlanGetCommand_initOptions, _PlannerPlanGetCommand_initValidators, _PlannerPlanGetCommand_initOptionSets, _PlannerPlanGetCommand_initTypes;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerPlanGetCommand extends GraphCommand {
    get name() {
        return commands.PLAN_GET;
    }
    get description() {
        return 'Get a Microsoft Planner plan';
    }
    constructor() {
        super();
        _PlannerPlanGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerPlanGetCommand_instances, "m", _PlannerPlanGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerPlanGetCommand_instances, "m", _PlannerPlanGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerPlanGetCommand_instances, "m", _PlannerPlanGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerPlanGetCommand_instances, "m", _PlannerPlanGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerPlanGetCommand_instances, "m", _PlannerPlanGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            let plan;
            if (args.options.id) {
                plan = await planner.getPlanById(args.options.id);
            }
            else if (args.options.rosterId) {
                plan = await planner.getPlanByRosterId(args.options.rosterId);
            }
            else {
                const groupId = await this.getGroupId(args);
                plan = await planner.getPlanByTitle(args.options.title, groupId);
            }
            const result = await this.getPlanDetails(plan);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getPlanDetails(plan) {
        const requestOptionsTaskDetails = {
            url: `${this.resource}/v1.0/planner/plans/${plan.id}/details`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                Prefer: 'return=representation'
            },
            responseType: 'json'
        };
        const planDetails = await request.get(requestOptionsTaskDetails);
        return { ...plan, ...planDetails };
    }
    async getGroupId(args) {
        if (args.options.ownerGroupId) {
            return args.options.ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName);
    }
}
_PlannerPlanGetCommand_instances = new WeakSet(), _PlannerPlanGetCommand_initTelemetry = function _PlannerPlanGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined'
        });
    });
}, _PlannerPlanGetCommand_initOptions = function _PlannerPlanGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '--rosterId [rosterId]'
    }, {
        option: '--ownerGroupId [ownerGroupId]'
    }, {
        option: '--ownerGroupName [ownerGroupName]'
    });
}, _PlannerPlanGetCommand_initValidators = function _PlannerPlanGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
        }
        return true;
    });
}, _PlannerPlanGetCommand_initOptionSets = function _PlannerPlanGetCommand_initOptionSets() {
    this.optionSets.push({
        options: ['id', 'title', 'rosterId']
    }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => {
            return args.options.title !== undefined;
        }
    });
}, _PlannerPlanGetCommand_initTypes = function _PlannerPlanGetCommand_initTypes() {
    this.types.string.push('id', 'title', 'ownerGroupId', 'ownerGroupName', 'rosterId');
};
export default new PlannerPlanGetCommand();
//# sourceMappingURL=plan-get.js.map