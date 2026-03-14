var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionEventTypeSetCommand_instances, _PurviewRetentionEventTypeSetCommand_initTelemetry, _PurviewRetentionEventTypeSetCommand_initOptions, _PurviewRetentionEventTypeSetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionEventTypeSetCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONEVENTTYPE_SET;
    }
    get description() {
        return 'Update a retention event type';
    }
    constructor() {
        super();
        _PurviewRetentionEventTypeSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeSetCommand_instances, "m", _PurviewRetentionEventTypeSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeSetCommand_instances, "m", _PurviewRetentionEventTypeSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeSetCommand_instances, "m", _PurviewRetentionEventTypeSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.log(`Updating retention event type with id ${args.options.id}`);
        }
        try {
            const requestBody = {
                description: args.options.description
            };
            const requestOptions = {
                url: `${this.resource}/v1.0/security/triggerTypes/retentionEventTypes/${args.options.id}`,
                headers: {
                    accept: 'application/json'
                },
                responseType: 'json',
                data: requestBody
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PurviewRetentionEventTypeSetCommand_instances = new WeakSet(), _PurviewRetentionEventTypeSetCommand_initTelemetry = function _PurviewRetentionEventTypeSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined'
        });
    });
}, _PurviewRetentionEventTypeSetCommand_initOptions = function _PurviewRetentionEventTypeSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-d, --description [description]'
    });
}, _PurviewRetentionEventTypeSetCommand_initValidators = function _PurviewRetentionEventTypeSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID.`;
        }
        if (!args.options.description) {
            return 'Specify at least one option to update.';
        }
        return true;
    });
};
export default new PurviewRetentionEventTypeSetCommand();
//# sourceMappingURL=retentioneventtype-set.js.map