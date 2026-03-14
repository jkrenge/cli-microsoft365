var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPageTemplateListCommand_instances, _SpoPageTemplateListCommand_initOptions, _SpoPageTemplateListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoPageTemplateListCommand extends SpoCommand {
    get name() {
        return commands.PAGE_TEMPLATE_LIST;
    }
    get description() {
        return 'Lists all page templates in the given site';
    }
    defaultProperties() {
        return ['Title', 'FileName', 'Id', 'PageLayoutType', 'Url'];
    }
    constructor() {
        super();
        _SpoPageTemplateListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPageTemplateListCommand_instances, "m", _SpoPageTemplateListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPageTemplateListCommand_instances, "m", _SpoPageTemplateListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving templates...`);
        }
        try {
            const res = await odata.getAllItems(`${args.options.webUrl}/_api/sitepages/pages/templates`);
            await logger.log(res);
        }
        catch (err) {
            // The API returns a 404 when no templates are created on the site collection
            if (err && err.response && err.response.status && err.response.status === 404) {
                await logger.log([]);
                return;
            }
            return this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoPageTemplateListCommand_instances = new WeakSet(), _SpoPageTemplateListCommand_initOptions = function _SpoPageTemplateListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoPageTemplateListCommand_initValidators = function _SpoPageTemplateListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoPageTemplateListCommand();
//# sourceMappingURL=page-template-list.js.map