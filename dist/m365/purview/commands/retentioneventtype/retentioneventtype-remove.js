var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionEventTypeRemoveCommand_instances, _PurviewRetentionEventTypeRemoveCommand_initTelemetry, _PurviewRetentionEventTypeRemoveCommand_initOptions, _PurviewRetentionEventTypeRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionEventTypeRemoveCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONEVENTTYPE_REMOVE;
    }
    get description() {
        return 'Delete a retention event type';
    }
    constructor() {
        super();
        _PurviewRetentionEventTypeRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeRemoveCommand_instances, "m", _PurviewRetentionEventTypeRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeRemoveCommand_instances, "m", _PurviewRetentionEventTypeRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeRemoveCommand_instances, "m", _PurviewRetentionEventTypeRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeRetentionEventType(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the retention event type with id ${args.options.id}?` });
            if (result) {
                await this.removeRetentionEventType(args.options);
            }
        }
    }
    async removeRetentionEventType(options) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/security/triggerTypes/retentionEventTypes/${options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PurviewRetentionEventTypeRemoveCommand_instances = new WeakSet(), _PurviewRetentionEventTypeRemoveCommand_initTelemetry = function _PurviewRetentionEventTypeRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _PurviewRetentionEventTypeRemoveCommand_initOptions = function _PurviewRetentionEventTypeRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _PurviewRetentionEventTypeRemoveCommand_initValidators = function _PurviewRetentionEventTypeRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID.`;
        }
        return true;
    });
};
export default new PurviewRetentionEventTypeRemoveCommand();
//# sourceMappingURL=retentioneventtype-remove.js.map