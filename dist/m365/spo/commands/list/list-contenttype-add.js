var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListContentTypeAddCommand_instances, _SpoListContentTypeAddCommand_initTelemetry, _SpoListContentTypeAddCommand_initOptions, _SpoListContentTypeAddCommand_initValidators, _SpoListContentTypeAddCommand_initTypes, _SpoListContentTypeAddCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListContentTypeAddCommand extends SpoCommand {
    get name() {
        return commands.LIST_CONTENTTYPE_ADD;
    }
    get description() {
        return 'Adds content type to list';
    }
    constructor() {
        super();
        _SpoListContentTypeAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListContentTypeAddCommand_instances, "m", _SpoListContentTypeAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeAddCommand_instances, "m", _SpoListContentTypeAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeAddCommand_instances, "m", _SpoListContentTypeAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeAddCommand_instances, "m", _SpoListContentTypeAddCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeAddCommand_instances, "m", _SpoListContentTypeAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            const list = (args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl);
            await logger.logToStderr(`Adding content type ${args.options.id} to list ${list} in site at ${args.options.webUrl}...`);
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
        const requestBody = {
            contentTypeId: args.options.id
        };
        const requestOptions = {
            url: `${requestUrl}/ContentTypes/AddAvailableContentType`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: requestBody,
            responseType: 'json'
        };
        try {
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListContentTypeAddCommand_instances = new WeakSet(), _SpoListContentTypeAddCommand_initTelemetry = function _SpoListContentTypeAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListContentTypeAddCommand_initOptions = function _SpoListContentTypeAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-i, --id <id>'
    });
}, _SpoListContentTypeAddCommand_initValidators = function _SpoListContentTypeAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoListContentTypeAddCommand_initTypes = function _SpoListContentTypeAddCommand_initTypes() {
    this.types.string.push('id', 'i');
}, _SpoListContentTypeAddCommand_initOptionSets = function _SpoListContentTypeAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListContentTypeAddCommand();
//# sourceMappingURL=list-contenttype-add.js.map