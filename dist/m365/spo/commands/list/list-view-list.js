var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListViewListCommand_instances, _SpoListViewListCommand_initTelemetry, _SpoListViewListCommand_initOptions, _SpoListViewListCommand_initValidators, _SpoListViewListCommand_initOptionSets;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListViewListCommand extends SpoCommand {
    get name() {
        return commands.LIST_VIEW_LIST;
    }
    get description() {
        return 'Lists views configured on the specified list';
    }
    defaultProperties() {
        return ['Id', 'Title', 'DefaultView', 'Hidden', 'BaseViewId'];
    }
    constructor() {
        super();
        _SpoListViewListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListViewListCommand_instances, "m", _SpoListViewListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListViewListCommand_instances, "m", _SpoListViewListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListViewListCommand_instances, "m", _SpoListViewListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListViewListCommand_instances, "m", _SpoListViewListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let requestUrl = `${args.options.webUrl}/_api/web`;
        if (args.options.listId) {
            requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/views`;
        }
        else if (args.options.listTitle) {
            requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/views`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/views`;
        }
        try {
            const res = await odata.getAllItems(requestUrl);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListViewListCommand_instances = new WeakSet(), _SpoListViewListCommand_initTelemetry = function _SpoListViewListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListViewListCommand_initOptions = function _SpoListViewListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListViewListCommand_initValidators = function _SpoListViewListCommand_initValidators() {
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
}, _SpoListViewListCommand_initOptionSets = function _SpoListViewListCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListViewListCommand();
//# sourceMappingURL=list-view-list.js.map