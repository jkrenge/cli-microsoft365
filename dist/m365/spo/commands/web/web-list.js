var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebListCommand_instances, _SpoWebListCommand_initOptions, _SpoWebListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoWebListCommand extends SpoCommand {
    get name() {
        return commands.WEB_LIST;
    }
    get description() {
        return 'Lists subsites of the specified site';
    }
    defaultProperties() {
        return ['Title', 'Url', 'Id'];
    }
    constructor() {
        super();
        _SpoWebListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebListCommand_instances, "m", _SpoWebListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebListCommand_instances, "m", _SpoWebListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving all webs in site at ${args.options.url}...`);
        }
        let requestUrl = `${args.options.url}/_api/web/webs`;
        if (args.options.output !== 'json') {
            requestUrl += '?$select=Title,Id,URL';
        }
        try {
            const webProperties = await odata.getAllItems(requestUrl);
            await logger.log(webProperties);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoWebListCommand_instances = new WeakSet(), _SpoWebListCommand_initOptions = function _SpoWebListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    });
}, _SpoWebListCommand_initValidators = function _SpoWebListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
};
export default new SpoWebListCommand();
//# sourceMappingURL=web-list.js.map