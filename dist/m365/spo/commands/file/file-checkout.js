var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileCheckoutCommand_instances, _SpoFileCheckoutCommand_initTelemetry, _SpoFileCheckoutCommand_initOptions, _SpoFileCheckoutCommand_initValidators, _SpoFileCheckoutCommand_initOptionSets, _SpoFileCheckoutCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileCheckoutCommand extends SpoCommand {
    get name() {
        return commands.FILE_CHECKOUT;
    }
    get description() {
        return 'Checks out specified file';
    }
    constructor() {
        super();
        _SpoFileCheckoutCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutCommand_instances, "m", _SpoFileCheckoutCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutCommand_instances, "m", _SpoFileCheckoutCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutCommand_instances, "m", _SpoFileCheckoutCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutCommand_instances, "m", _SpoFileCheckoutCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileCheckoutCommand_instances, "m", _SpoFileCheckoutCommand_initTypes).call(this);
    }
    getExcludedOptionsWithUrls() {
        return ['url'];
    }
    async commandAction(logger, args) {
        let requestUrl = '';
        if (args.options.id) {
            requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${formatting.encodeQueryParameter(args.options.id)}')/checkout`;
        }
        if (args.options.url) {
            const serverRelativePath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.url);
            requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/checkout`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFileCheckoutCommand_instances = new WeakSet(), _SpoFileCheckoutCommand_initTelemetry = function _SpoFileCheckoutCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            url: typeof args.options.url !== 'undefined'
        });
    });
}, _SpoFileCheckoutCommand_initOptions = function _SpoFileCheckoutCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--url [url]'
    }, {
        option: '-i, --id [id]'
    });
}, _SpoFileCheckoutCommand_initValidators = function _SpoFileCheckoutCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.id) {
            if (!validation.isValidGuid(args.options.id)) {
                return `${args.options.id} is not a valid GUID`;
            }
        }
        return true;
    });
}, _SpoFileCheckoutCommand_initOptionSets = function _SpoFileCheckoutCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'url'] });
}, _SpoFileCheckoutCommand_initTypes = function _SpoFileCheckoutCommand_initTypes() {
    this.types.string.push('webUrl', 'url', 'id');
};
export default new SpoFileCheckoutCommand();
//# sourceMappingURL=file-checkout.js.map