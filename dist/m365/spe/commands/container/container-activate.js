var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpeContainerActivateCommand_instances, _SpeContainerActivateCommand_initOptions, _SpeContainerActivateCommand_initTypes;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
class SpeContainerActivateCommand extends GraphCommand {
    get name() {
        return commands.CONTAINER_ACTIVATE;
    }
    get description() {
        return 'Activates a container';
    }
    constructor() {
        super();
        _SpeContainerActivateCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpeContainerActivateCommand_instances, "m", _SpeContainerActivateCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpeContainerActivateCommand_instances, "m", _SpeContainerActivateCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Activating a container with id '${args.options.id}'...`);
        }
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(args.options.id)}/activate`,
                headers: {
                    'content-type': 'application/json;odata=nometadata',
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpeContainerActivateCommand_instances = new WeakSet(), _SpeContainerActivateCommand_initOptions = function _SpeContainerActivateCommand_initOptions() {
    this.options.unshift({ option: '-i, --id <id>' });
}, _SpeContainerActivateCommand_initTypes = function _SpeContainerActivateCommand_initTypes() {
    this.types.string.push('id');
};
export default new SpeContainerActivateCommand();
//# sourceMappingURL=container-activate.js.map