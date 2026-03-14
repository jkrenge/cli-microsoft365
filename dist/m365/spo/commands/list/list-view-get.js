var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListViewGetCommand_instances, _SpoListViewGetCommand_initTelemetry, _SpoListViewGetCommand_initOptions, _SpoListViewGetCommand_initValidators, _SpoListViewGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListViewGetCommand extends SpoCommand {
    get name() {
        return commands.LIST_VIEW_GET;
    }
    get description() {
        return 'Gets information about specific list view';
    }
    constructor() {
        super();
        _SpoListViewGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListViewGetCommand_instances, "m", _SpoListViewGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListViewGetCommand_instances, "m", _SpoListViewGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListViewGetCommand_instances, "m", _SpoListViewGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListViewGetCommand_instances, "m", _SpoListViewGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const baseRestUrl = `${args.options.webUrl}/_api/web`;
        let listRestUrl = '';
        if (args.options.listId) {
            listRestUrl = `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            listRestUrl = `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            listRestUrl = `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        const viewRestUrl = `/views/${(args.options.id ? `getById('${formatting.encodeQueryParameter(args.options.id)}')` : `getByTitle('${formatting.encodeQueryParameter(args.options.title)}')`)}`;
        const requestOptions = {
            url: `${baseRestUrl}${listRestUrl}${viewRestUrl}`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const result = await request.get(requestOptions);
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListViewGetCommand_instances = new WeakSet(), _SpoListViewGetCommand_initTelemetry = function _SpoListViewGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined'
        });
    });
}, _SpoListViewGetCommand_initOptions = function _SpoListViewGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '--id [id]'
    }, {
        option: '--title [title]'
    });
}, _SpoListViewGetCommand_initValidators = function _SpoListViewGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        if (args.options.id &&
            !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} in option id is not a valid GUID`;
        }
        return true;
    });
}, _SpoListViewGetCommand_initOptionSets = function _SpoListViewGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['id', 'title'] });
};
export default new SpoListViewGetCommand();
//# sourceMappingURL=list-view-get.js.map