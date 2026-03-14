var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoContentTypeListCommand_instances, _SpoContentTypeListCommand_initTelemetry, _SpoContentTypeListCommand_initOptions, _SpoContentTypeListCommand_initValidators, _SpoContentTypeListCommand_initTypes;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoContentTypeListCommand extends SpoCommand {
    get name() {
        return commands.CONTENTTYPE_LIST;
    }
    get description() {
        return 'Lists all available content types in the specified site';
    }
    defaultProperties() {
        return ['StringId', 'Name', 'Hidden', 'ReadOnly', 'Sealed'];
    }
    constructor() {
        super();
        _SpoContentTypeListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoContentTypeListCommand_instances, "m", _SpoContentTypeListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeListCommand_instances, "m", _SpoContentTypeListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeListCommand_instances, "m", _SpoContentTypeListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeListCommand_instances, "m", _SpoContentTypeListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            let requestUrl = `${args.options.webUrl}/_api/web/ContentTypes?$expand=Parent`;
            if (args.options.category) {
                requestUrl += `&$filter=Group eq '${formatting.encodeQueryParameter(args.options.category)}'`;
            }
            const res = await odata.getAllItems(requestUrl);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoContentTypeListCommand_instances = new WeakSet(), _SpoContentTypeListCommand_initTelemetry = function _SpoContentTypeListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            category: typeof args.options.category !== 'undefined'
        });
    });
}, _SpoContentTypeListCommand_initOptions = function _SpoContentTypeListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-c, --category [category]'
    });
}, _SpoContentTypeListCommand_initValidators = function _SpoContentTypeListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
}, _SpoContentTypeListCommand_initTypes = function _SpoContentTypeListCommand_initTypes() {
    this.types.string.push('webUrl', 'category');
};
export default new SpoContentTypeListCommand();
//# sourceMappingURL=contenttype-list.js.map