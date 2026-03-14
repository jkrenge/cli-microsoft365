var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageMessageRemoveCommand_instances, _VivaEngageMessageRemoveCommand_initTelemetry, _VivaEngageMessageRemoveCommand_initOptions, _VivaEngageMessageRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
class VivaEngageMessageRemoveCommand extends VivaEngageCommand {
    get name() {
        return commands.ENGAGE_MESSAGE_REMOVE;
    }
    get description() {
        return 'Removes a Viva Engage message';
    }
    constructor() {
        super();
        _VivaEngageMessageRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _VivaEngageMessageRemoveCommand_instances, "m", _VivaEngageMessageRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageRemoveCommand_instances, "m", _VivaEngageMessageRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageRemoveCommand_instances, "m", _VivaEngageMessageRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeMessage(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Viva Engage message ${args.options.id}?` });
            if (result) {
                await this.removeMessage(args.options);
            }
        }
    }
    async removeMessage(options) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1/messages/${options.id}.json`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json;odata=nometadata'
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
_VivaEngageMessageRemoveCommand_instances = new WeakSet(), _VivaEngageMessageRemoveCommand_initTelemetry = function _VivaEngageMessageRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _VivaEngageMessageRemoveCommand_initOptions = function _VivaEngageMessageRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--id <id>'
    }, {
        option: '-f, --force'
    });
}, _VivaEngageMessageRemoveCommand_initValidators = function _VivaEngageMessageRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (typeof args.options.id !== 'number') {
            return `${args.options.id} is not a number`;
        }
        return true;
    });
};
export default new VivaEngageMessageRemoveCommand();
//# sourceMappingURL=engage-message-remove.js.map