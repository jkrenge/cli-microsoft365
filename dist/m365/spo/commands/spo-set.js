var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSetCommand_instances, _SpoSetCommand_initOptions, _SpoSetCommand_initValidators, _SpoSetCommand_initTypes;
import auth from '../../../Auth.js';
import { urlUtil } from '../../../utils/urlUtil.js';
import { validation } from '../../../utils/validation.js';
import SpoCommand from '../../base/SpoCommand.js';
import commands from '../commands.js';
class SpoSetCommand extends SpoCommand {
    get name() {
        return commands.SET;
    }
    get description() {
        return 'Sets the URL of the root SharePoint site collection for use in SPO commands';
    }
    constructor() {
        super();
        _SpoSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSetCommand_instances, "m", _SpoSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSetCommand_instances, "m", _SpoSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSetCommand_instances, "m", _SpoSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        auth.connection.spoUrl = urlUtil.removeTrailingSlashes(args.options.url);
        try {
            await auth.storeConnectionInfo();
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSetCommand_instances = new WeakSet(), _SpoSetCommand_initOptions = function _SpoSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    });
}, _SpoSetCommand_initValidators = function _SpoSetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
}, _SpoSetCommand_initTypes = function _SpoSetCommand_initTypes() {
    this.types.string.push('url');
};
export default new SpoSetCommand();
//# sourceMappingURL=spo-set.js.map