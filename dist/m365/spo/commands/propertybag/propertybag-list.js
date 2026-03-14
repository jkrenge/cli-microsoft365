var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPropertyBagListCommand_instances, _SpoPropertyBagListCommand_initTelemetry, _SpoPropertyBagListCommand_initOptions, _SpoPropertyBagListCommand_initValidators;
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoPropertyBagBaseCommand } from './propertybag-base.js';
class SpoPropertyBagListCommand extends SpoPropertyBagBaseCommand {
    get name() {
        return commands.PROPERTYBAG_LIST;
    }
    get description() {
        return 'Gets property bag values';
    }
    constructor() {
        super();
        _SpoPropertyBagListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPropertyBagListCommand_instances, "m", _SpoPropertyBagListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoPropertyBagListCommand_instances, "m", _SpoPropertyBagListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPropertyBagListCommand_instances, "m", _SpoPropertyBagListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const contextResponse = await spo.getRequestDigest(args.options.webUrl);
            this.formDigestValue = contextResponse.FormDigestValue;
            const identityResp = await spo.getCurrentWebIdentity(args.options.webUrl, this.formDigestValue);
            let propertyBagData;
            const opts = args.options;
            if (opts.folder) {
                propertyBagData = await this.getFolderPropertyBag(identityResp, opts.webUrl, opts.folder, logger);
            }
            else {
                propertyBagData = await this.getWebPropertyBag(identityResp, opts.webUrl, logger);
            }
            await logger.log(this.formatOutput(propertyBagData));
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    /**
     * The property bag data returned from the client.svc/ProcessQuery response
     * has to be formatted before displayed since the key, value objects
     * carry extra information.
     * @param propertyBag client.svc property bag javascript object
     */
    formatOutput(propertyBag) {
        const result = [];
        const keys = Object.keys(propertyBag);
        for (let i = 0; i < keys.length; i++) {
            if (keys[i] === '_ObjectType_') {
                // this is system data, do not include it
                continue;
            }
            const formattedProp = this.formatProperty(keys[i], propertyBag[keys[i]]);
            result.push(formattedProp);
        }
        return result;
    }
}
_SpoPropertyBagListCommand_instances = new WeakSet(), _SpoPropertyBagListCommand_initTelemetry = function _SpoPropertyBagListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folder: typeof args.options.folder !== 'undefined'
        });
    });
}, _SpoPropertyBagListCommand_initOptions = function _SpoPropertyBagListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--folder [folder]'
    });
}, _SpoPropertyBagListCommand_initValidators = function _SpoPropertyBagListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoPropertyBagListCommand();
//# sourceMappingURL=propertybag-list.js.map