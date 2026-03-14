var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageMessageGetCommand_instances, _VivaEngageMessageGetCommand_initOptions, _VivaEngageMessageGetCommand_initValidators;
import request from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
class VivaEngageMessageGetCommand extends VivaEngageCommand {
    get name() {
        return commands.ENGAGE_MESSAGE_GET;
    }
    get description() {
        return 'Returns a Viva Engage message';
    }
    constructor() {
        super();
        _VivaEngageMessageGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _VivaEngageMessageGetCommand_instances, "m", _VivaEngageMessageGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _VivaEngageMessageGetCommand_instances, "m", _VivaEngageMessageGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/v1/messages/${args.options.id}.json`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
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
_VivaEngageMessageGetCommand_instances = new WeakSet(), _VivaEngageMessageGetCommand_initOptions = function _VivaEngageMessageGetCommand_initOptions() {
    this.options.unshift({
        option: '--id <id>'
    });
}, _VivaEngageMessageGetCommand_initValidators = function _VivaEngageMessageGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (typeof args.options.id !== 'number') {
            return `${args.options.id} is not a number`;
        }
        return true;
    });
};
export default new VivaEngageMessageGetCommand();
//# sourceMappingURL=engage-message-get.js.map