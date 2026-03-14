var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerBucketRemoveCommand_instances, _PlannerBucketRemoveCommand_initTelemetry, _PlannerBucketRemoveCommand_initOptions, _PlannerBucketRemoveCommand_initValidators, _PlannerBucketRemoveCommand_initOptionSets, _PlannerBucketRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { planner } from '../../../../utils/planner.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerBucketRemoveCommand extends GraphCommand {
    get name() {
        return commands.BUCKET_REMOVE;
    }
    get description() {
        return 'Removes the Microsoft Planner bucket from a plan';
    }
    constructor() {
        super();
        _PlannerBucketRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerBucketRemoveCommand_instances, "m", _PlannerBucketRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerBucketRemoveCommand_instances, "m", _PlannerBucketRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerBucketRemoveCommand_instances, "m", _PlannerBucketRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerBucketRemoveCommand_instances, "m", _PlannerBucketRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerBucketRemoveCommand_instances, "m", _PlannerBucketRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeBucket = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing bucket...`);
            }
            try {
                const bucket = await this.getBucket(args);
                const requestOptions = {
                    url: `${this.resource}/v1.0/planner/buckets/${bucket.id}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none',
                        'if-match': bucket['@odata.etag']
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
            await removeBucket();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the bucket ${args.options.id || args.options.name}?` });
            if (result) {
                await removeBucket();
            }
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
_PlannerBucketRemoveCommand_instances = new WeakSet(), _PlannerBucketRemoveCommand_initTelemetry = function _PlannerBucketRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            planId: typeof args.options.planId !== 'undefined',
            planTitle: typeof args.options.planTitle !== 'undefined',
            rosterId: typeof args.options.rosterId !== 'undefined',
            ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
            ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
            force: args.options.force || false
        });
    });
}, _PlannerBucketRemoveCommand_initOptions = function _PlannerBucketRemoveCommand_initOptions() {
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
        option: '-f, --force'
    });
}, _PlannerBucketRemoveCommand_initValidators = function _PlannerBucketRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id) {
            if (args.options.planId || args.options.planTitle || args.options.rosterId || args.options.ownerGroupId || args.options.ownerGroupName) {
                return 'Don\'t specify planId, planTitle, rosterId, ownerGroupId or ownerGroupName when using id';
            }
        }
        else {
            if (args.options.planTitle) {
                if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
                    return `${args.options.ownerGroupId} is not a valid GUID`;
                }
            }
            else if (args.options.planId) {
                if (args.options.ownerGroupId || args.options.ownerGroupName) {
                    return 'Don\'t specify ownerGroupId or ownerGroupName when using planId';
                }
            }
            else {
                if (args.options.ownerGroupId || args.options.ownerGroupName) {
                    return 'Don\'t specify ownerGroupId or ownerGroupName when using rosterId';
                }
            }
        }
        return true;
    });
}, _PlannerBucketRemoveCommand_initOptionSets = function _PlannerBucketRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] }, {
        options: ['planId', 'planTitle', 'rosterId'],
        runsWhen: (args) => args.options.name !== undefined
    }, {
        options: ['ownerGroupId', 'ownerGroupName'],
        runsWhen: (args) => args.options.planTitle !== undefined
    });
}, _PlannerBucketRemoveCommand_initTypes = function _PlannerBucketRemoveCommand_initTypes() {
    this.types.string.push('id', 'name', 'planId', 'planTitle', 'ownerGroupId', 'ownerGroupName', 'rosterId ');
};
export default new PlannerBucketRemoveCommand();
//# sourceMappingURL=bucket-remove.js.map