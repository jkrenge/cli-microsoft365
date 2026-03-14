var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpeContainerGetCommand_instances, _SpeContainerGetCommand_initOptions, _SpeContainerGetCommand_initTypes;
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
class SpeContainerGetCommand extends GraphCommand {
    get name() {
        return commands.CONTAINER_GET;
    }
    get description() {
        return 'Gets a container of a specific container type';
    }
    constructor() {
        super();
        _SpeContainerGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpeContainerGetCommand_instances, "m", _SpeContainerGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpeContainerGetCommand_instances, "m", _SpeContainerGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Getting a container with id '${args.options.id}'...`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/storage/fileStorage/containers/${args.options.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpeContainerGetCommand_instances = new WeakSet(), _SpeContainerGetCommand_initOptions = function _SpeContainerGetCommand_initOptions() {
    this.options.unshift({ option: '-i, --id <id>' });
}, _SpeContainerGetCommand_initTypes = function _SpeContainerGetCommand_initTypes() {
    this.types.string.push('id');
};
export default new SpeContainerGetCommand();
//# sourceMappingURL=container-get.js.map