var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionEventRemoveCommand_instances, _PurviewRetentionEventRemoveCommand_initTelemetry, _PurviewRetentionEventRemoveCommand_initOptions, _PurviewRetentionEventRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionEventRemoveCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONEVENT_REMOVE;
    }
    get description() {
        return 'Delete a retention event';
    }
    constructor() {
        super();
        _PurviewRetentionEventRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventRemoveCommand_instances, "m", _PurviewRetentionEventRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventRemoveCommand_instances, "m", _PurviewRetentionEventRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventRemoveCommand_instances, "m", _PurviewRetentionEventRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeRetentionEvent(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the retention event ${args.options.id}?` });
            if (result) {
                await this.removeRetentionEvent(args.options);
            }
        }
    }
    async removeRetentionEvent(options) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/security/triggers/retentionEvents/${options.id}`,
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
_PurviewRetentionEventRemoveCommand_instances = new WeakSet(), _PurviewRetentionEventRemoveCommand_initTelemetry = function _PurviewRetentionEventRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _PurviewRetentionEventRemoveCommand_initOptions = function _PurviewRetentionEventRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _PurviewRetentionEventRemoveCommand_initValidators = function _PurviewRetentionEventRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID.`;
        }
        return true;
    });
};
export default new PurviewRetentionEventRemoveCommand();
//# sourceMappingURL=retentionevent-remove.js.map