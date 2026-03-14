var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPropertyBagSetCommand_instances, _SpoPropertyBagSetCommand_initTelemetry, _SpoPropertyBagSetCommand_initOptions, _SpoPropertyBagSetCommand_initValidators;
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoPropertyBagBaseCommand } from './propertybag-base.js';
class SpoPropertyBagSetCommand extends SpoPropertyBagBaseCommand {
    get name() {
        return commands.PROPERTYBAG_SET;
    }
    get description() {
        return 'Sets the value of the specified property in the property bag. Adds the property if it does not exist';
    }
    constructor() {
        super();
        _SpoPropertyBagSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPropertyBagSetCommand_instances, "m", _SpoPropertyBagSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoPropertyBagSetCommand_instances, "m", _SpoPropertyBagSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPropertyBagSetCommand_instances, "m", _SpoPropertyBagSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const contextResponse = await spo.getRequestDigest(args.options.webUrl);
            this.formDigestValue = contextResponse.FormDigestValue;
            let identityResp = await spo.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
            const webIdentityResp = identityResp;
            // Check if web no script enabled or not
            // Cannot set property bag value if no script is enabled
            const isNoScriptSite = await this.isNoScriptSite(identityResp, args.options, logger);
            if (isNoScriptSite) {
                throw 'Site has NoScript enabled, and setting property bag values is not supported';
            }
            const opts = args.options;
            if (opts.folder) {
                // get the folder guid instead of the web guid
                identityResp = await spo.getFolderIdentity(webIdentityResp.objectIdentity, opts.webUrl, opts.folder, this.formDigestValue);
            }
            await this.setProperty(identityResp, args.options, logger);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    setProperty(identityResp, options, logger) {
        return SpoPropertyBagBaseCommand.setProperty(options.key, options.value, options.webUrl, this.formDigestValue, identityResp, logger, this.debug, options.folder);
    }
    isNoScriptSite(webIdentityResp, options, logger) {
        return SpoPropertyBagBaseCommand.isNoScriptSite(options.webUrl, this.formDigestValue, webIdentityResp, logger, this.debug);
    }
}
_SpoPropertyBagSetCommand_instances = new WeakSet(), _SpoPropertyBagSetCommand_initTelemetry = function _SpoPropertyBagSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folder: typeof args.options.folder !== 'undefined'
        });
    });
}, _SpoPropertyBagSetCommand_initOptions = function _SpoPropertyBagSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-k, --key <key>'
    }, {
        option: '-v, --value <value>'
    }, {
        option: '--folder [folder]'
    });
}, _SpoPropertyBagSetCommand_initValidators = function _SpoPropertyBagSetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoPropertyBagSetCommand();
//# sourceMappingURL=propertybag-set.js.map