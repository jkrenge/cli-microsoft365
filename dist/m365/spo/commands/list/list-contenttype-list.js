var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListContentTypeListCommand_instances, _SpoListContentTypeListCommand_initTelemetry, _SpoListContentTypeListCommand_initOptions, _SpoListContentTypeListCommand_initValidators, _SpoListContentTypeListCommand_initOptionSets, _SpoListContentTypeListCommand_initTypes;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListContentTypeListCommand extends SpoCommand {
    get name() {
        return commands.LIST_CONTENTTYPE_LIST;
    }
    get description() {
        return 'Lists content types configured on the list';
    }
    defaultProperties() {
        return ['StringId', 'Name', 'Hidden', 'ReadOnly', 'Sealed'];
    }
    constructor() {
        super();
        _SpoListContentTypeListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListContentTypeListCommand_instances, "m", _SpoListContentTypeListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeListCommand_instances, "m", _SpoListContentTypeListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeListCommand_instances, "m", _SpoListContentTypeListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeListCommand_instances, "m", _SpoListContentTypeListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeListCommand_instances, "m", _SpoListContentTypeListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            const list = (args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl);
            await logger.logToStderr(`Retrieving content types information for list ${list} in site at ${args.options.webUrl}...`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.listId) {
            requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        try {
            const res = await odata.getAllItems(`${requestUrl}/ContentTypes?$expand=Parent`);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListContentTypeListCommand_instances = new WeakSet(), _SpoListContentTypeListCommand_initTelemetry = function _SpoListContentTypeListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListContentTypeListCommand_initOptions = function _SpoListContentTypeListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListContentTypeListCommand_initValidators = function _SpoListContentTypeListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId) {
            if (!validation.isValidGuid(args.options.listId)) {
                return `${args.options.listId} is not a valid GUID`;
            }
        }
        return true;
    });
}, _SpoListContentTypeListCommand_initOptionSets = function _SpoListContentTypeListCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
}, _SpoListContentTypeListCommand_initTypes = function _SpoListContentTypeListCommand_initTypes() {
    this.types.string.push('webUrl', 'listId', 'listTitle', 'listUrl');
};
export default new SpoListContentTypeListCommand();
//# sourceMappingURL=list-contenttype-list.js.map