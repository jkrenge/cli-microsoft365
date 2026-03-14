var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListViewSetCommand_instances, _SpoListViewSetCommand_initTelemetry, _SpoListViewSetCommand_initOptions, _SpoListViewSetCommand_initValidators, _SpoListViewSetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListViewSetCommand extends SpoCommand {
    allowUnknownOptions() {
        return true;
    }
    get name() {
        return commands.LIST_VIEW_SET;
    }
    get description() {
        return 'Updates existing list view';
    }
    constructor() {
        super();
        _SpoListViewSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListViewSetCommand_instances, "m", _SpoListViewSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListViewSetCommand_instances, "m", _SpoListViewSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListViewSetCommand_instances, "m", _SpoListViewSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListViewSetCommand_instances, "m", _SpoListViewSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let listRestUrl = '';
        if (args.options.listId) {
            listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            listRestUrl = `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            listRestUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        const viewRestUrl = `/views/${(args.options.id ? `GetById('${formatting.encodeQueryParameter(args.options.id)}')` : `GetByTitle('${formatting.encodeQueryParameter(args.options.title)}')`)}`;
        try {
            const res = await spo.getRequestDigest(args.options.webUrl);
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/${listRestUrl}${viewRestUrl}`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue,
                    'content-type': 'application/json;odata=nometadata',
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: this.getPayload(args.options)
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getPayload(options) {
        const payload = {};
        const excludeOptions = [
            'webUrl',
            'listId',
            'listTitle',
            'listUrl',
            'id',
            'title',
            'debug',
            'verbose',
            'output'
        ];
        Object.keys(options).forEach(key => {
            if (excludeOptions.indexOf(key) === -1) {
                payload[key] = options[key];
            }
        });
        return payload;
    }
}
_SpoListViewSetCommand_instances = new WeakSet(), _SpoListViewSetCommand_initTelemetry = function _SpoListViewSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined'
        });
    });
}, _SpoListViewSetCommand_initOptions = function _SpoListViewSetCommand_initOptions() {
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
}, _SpoListViewSetCommand_initValidators = function _SpoListViewSetCommand_initValidators() {
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
}, _SpoListViewSetCommand_initOptionSets = function _SpoListViewSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['id', 'title'] });
};
export default new SpoListViewSetCommand();
//# sourceMappingURL=list-view-set.js.map