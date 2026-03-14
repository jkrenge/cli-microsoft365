var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerPlanRemoveCommand_instances, _PlannerPlanRemoveCommand_initTelemetry, _PlannerPlanRemoveCommand_initOptions, _PlannerPlanRemoveCommand_initValidators, _PlannerPlanRemoveCommand_initOptionSets, _PlannerPlanRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerPlanRemoveCommand extends GraphCommand {
    get name() {
        return commands.PLAN_REMOVE;
    }
    get description() {
        return 'Removes the Microsoft Planner plan';
    }
    constructor() {
        super();
        _PlannerPlanRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerPlanRemoveCommand_instances, "m", _PlannerPlanRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerPlanRemoveCommand_instances, "m", _PlannerPlanRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerPlanRemoveCommand_instances, "m", _PlannerPlanRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerPlanRemoveCommand_instances, "m", _PlannerPlanRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerPlanRemoveCommand_instances, "m", _PlannerPlanRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removePlan = async () => {
            try {
                const plan = await this.getPlan(args);
                if (this.verbose) {
                    await logger.logToStderr(`Removing plan '${plan.title}' ...`);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/planner/plans/${plan.id}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none',
                        'if-match': plan['@odata.etag']
                    },
                    responseType: 'json'
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removePlan();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the plan ${args.options.id || args.options.title}?` });
            if (result) {
                await removePlan();
            }
        }
    }
    async getPlan(args) {
        const { id, title } = args.options;
        if (id) {
            return planner.getPlanById(id, 'minimal');
        }
        const groupId = await this.getGroupId(args);
        return planner.getPlanByTitle(title, groupId, 'minimal');
    }
    async getGroupId(args) {
        const { ownerGroupId, ownerGroupName } = args.options;
        if (ownerGroupId) {
            return ownerGroupId;
        }
        return entraGroup.getGroupIdByDisplayName(ownerGroupName);
    }
}
_PlannerPlanRemoveCommand_instances = new WeakSet(), _PlannerPlanRemoveCommand_initTelemetry = function _PlannerPlanRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _PlannerPlanRemoveCommand_initOptions = function _PlannerPlanRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '--ownerGroupId [ownerGroupId]'
    }, {
        option: '--ownerGroupName [ownerGroupName]'
    }, {
        option: '-f, --force'
    });
}, _PlannerPlanRemoveCommand_initValidators = function _PlannerPlanRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.title) {
            if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
                return 'Specify either ownerGroupId or ownerGroupName';
            }
            if (args.options.ownerGroupId && args.options.ownerGroupName) {
                return 'Specify either ownerGroupId or ownerGroupName but not both';
            }
            if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
                return `${args.options.ownerGroupId} is not a valid GUID`;
            }
        }
        else if (args.options.ownerGroupId || args.options.ownerGroupName) {
            return 'Don\'t specify ownerGroupId or ownerGroupName when using id';
        }
        return true;
    });
}, _PlannerPlanRemoveCommand_initOptionSets = function _PlannerPlanRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title'] });
}, _PlannerPlanRemoveCommand_initTypes = function _PlannerPlanRemoveCommand_initTypes() {
    this.types.string.push('id', 'title', 'ownerGroupId', 'ownerGroupName');
};
export default new PlannerPlanRemoveCommand();
//# sourceMappingURL=plan-remove.js.map