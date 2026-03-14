var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoPropertyBagGetCommand_instances, _SpoPropertyBagGetCommand_initTelemetry, _SpoPropertyBagGetCommand_initOptions, _SpoPropertyBagGetCommand_initValidators;
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoPropertyBagBaseCommand } from './propertybag-base.js';
class SpoPropertyBagGetCommand extends SpoPropertyBagBaseCommand {
    get name() {
        return commands.PROPERTYBAG_GET;
    }
    get description() {
        return 'Gets the value of the specified property from the property bag';
    }
    constructor() {
        super();
        _SpoPropertyBagGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoPropertyBagGetCommand_instances, "m", _SpoPropertyBagGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoPropertyBagGetCommand_instances, "m", _SpoPropertyBagGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoPropertyBagGetCommand_instances, "m", _SpoPropertyBagGetCommand_initValidators).call(this);
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
            const property = this.filterByKey(propertyBagData, args.options.key);
            if (property) {
                await logger.log(property.value);
            }
            else if (this.verbose) {
                await logger.logToStderr('Property not found.');
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    filterByKey(propertyBag, key) {
        const keys = Object.keys(propertyBag);
        for (let i = 0; i < keys.length; i++) {
            // we have to normalize the keys and values before we can filter
            // since they carry extra information
            // ex. : 'vti_level$  Int32' instead of 'vti_level'
            const formattedProperty = this.formatProperty(keys[i], propertyBag[keys[i]]);
            if (formattedProperty.key === key) {
                return formattedProperty;
            }
        }
        return null;
    }
}
_SpoPropertyBagGetCommand_instances = new WeakSet(), _SpoPropertyBagGetCommand_initTelemetry = function _SpoPropertyBagGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folder: typeof args.options.folder !== 'undefined'
        });
    });
}, _SpoPropertyBagGetCommand_initOptions = function _SpoPropertyBagGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-k, --key <key>'
    }, {
        option: '--folder [folder]'
    });
}, _SpoPropertyBagGetCommand_initValidators = function _SpoPropertyBagGetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoPropertyBagGetCommand();
//# sourceMappingURL=propertybag-get.js.map