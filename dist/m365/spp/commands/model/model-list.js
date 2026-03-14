var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SppModelListCommand_instances, _SppModelListCommand_initOptions, _SppModelListCommand_initTypes, _SppModelListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { spp } from '../../../../utils/spp.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SppModelListCommand extends SpoCommand {
    get name() {
        return commands.MODEL_LIST;
    }
    get description() {
        return 'Retrieves the list of unstructured document processing models';
    }
    defaultProperties() {
        return ['AIBuilderHybridModelType', 'ContentTypeName', 'LastTrained', 'UniqueId'];
    }
    constructor() {
        super();
        _SppModelListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SppModelListCommand_instances, "m", _SppModelListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SppModelListCommand_instances, "m", _SppModelListCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SppModelListCommand_instances, "m", _SppModelListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.log(`Retrieving models from ${args.options.siteUrl}...`);
            }
            const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
            await spp.assertSiteIsContentCenter(siteUrl, logger, this.verbose);
            const result = await odata.getAllItems(`${siteUrl}/_api/machinelearning/models`);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SppModelListCommand_instances = new WeakSet(), _SppModelListCommand_initOptions = function _SppModelListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    });
}, _SppModelListCommand_initTypes = function _SppModelListCommand_initTypes() {
    this.types.string.push('siteUrl');
}, _SppModelListCommand_initValidators = function _SppModelListCommand_initValidators() {
    this.validators.push(async (args) => {
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
};
export default new SppModelListCommand();
//# sourceMappingURL=model-list.js.map