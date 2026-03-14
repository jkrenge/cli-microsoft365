var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteRecycleBinItemListCommand_instances, _a, _SpoSiteRecycleBinItemListCommand_initTelemetry, _SpoSiteRecycleBinItemListCommand_initOptions, _SpoSiteRecycleBinItemListCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteRecycleBinItemListCommand extends SpoCommand {
    get name() {
        return commands.SITE_RECYCLEBINITEM_LIST;
    }
    get description() {
        return 'Lists items from recycle bin';
    }
    defaultProperties() {
        return ['Id', 'Title', 'DirName'];
    }
    constructor() {
        super();
        _SpoSiteRecycleBinItemListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemListCommand_instances, "m", _SpoSiteRecycleBinItemListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemListCommand_instances, "m", _SpoSiteRecycleBinItemListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteRecycleBinItemListCommand_instances, "m", _SpoSiteRecycleBinItemListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving all items from recycle bin at ${args.options.siteUrl}...`);
        }
        const state = args.options.secondary ? '2' : '1';
        let requestUrl = `${args.options.siteUrl}/_api/site/RecycleBin?$filter=(ItemState eq ${state})`;
        if (typeof args.options.type !== 'undefined') {
            const type = _a.recycleBinItemType.find(item => item.value === args.options.type);
            if (typeof type !== 'undefined') {
                requestUrl += ` and (ItemType eq ${type.id})`;
            }
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const response = await request.get(requestOptions);
            await logger.log(response.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = SpoSiteRecycleBinItemListCommand, _SpoSiteRecycleBinItemListCommand_instances = new WeakSet(), _SpoSiteRecycleBinItemListCommand_initTelemetry = function _SpoSiteRecycleBinItemListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            type: args.options.type,
            secondary: args.options.secondary
        });
    });
}, _SpoSiteRecycleBinItemListCommand_initOptions = function _SpoSiteRecycleBinItemListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-t, --type [type]',
        autocomplete: _a.recycleBinItemType.map(item => item.value)
    }, {
        option: '-s, --secondary'
    });
}, _SpoSiteRecycleBinItemListCommand_initValidators = function _SpoSiteRecycleBinItemListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (typeof args.options.type !== 'undefined' &&
            !_a.recycleBinItemType.some(item => item.value === args.options.type)) {
            return `${args.options.type} is not a valid value. Allowed values are ${_a.recycleBinItemType.map(item => item.value).join(', ')}`;
        }
        return true;
    });
};
SpoSiteRecycleBinItemListCommand.recycleBinItemType = [
    { id: 1, value: 'files' },
    { id: 3, value: 'listItems' },
    { id: 5, value: 'folders' }
];
export default new SpoSiteRecycleBinItemListCommand();
//# sourceMappingURL=site-recyclebinitem-list.js.map