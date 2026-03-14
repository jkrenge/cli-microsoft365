var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoContentTypeGetCommand_instances, _SpoContentTypeGetCommand_initTelemetry, _SpoContentTypeGetCommand_initOptions, _SpoContentTypeGetCommand_initValidators, _SpoContentTypeGetCommand_initTypes, _SpoContentTypeGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoContentTypeGetCommand extends SpoCommand {
    get name() {
        return commands.CONTENTTYPE_GET;
    }
    get description() {
        return 'Retrieves information about the specified list or site content type';
    }
    constructor() {
        super();
        _SpoContentTypeGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoContentTypeGetCommand_instances, "m", _SpoContentTypeGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeGetCommand_instances, "m", _SpoContentTypeGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeGetCommand_instances, "m", _SpoContentTypeGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeGetCommand_instances, "m", _SpoContentTypeGetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeGetCommand_instances, "m", _SpoContentTypeGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let requestUrl = `${args.options.webUrl}/_api/web`;
        if (args.options.listId) {
            requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        requestUrl += "/contenttypes";
        if (args.options.id) {
            requestUrl += `('${formatting.encodeQueryParameter(args.options.id)}')?$expand=Parent`;
        }
        else if (args.options.name) {
            requestUrl += `?$filter=Name eq '${formatting.encodeQueryParameter(args.options.name)}'&$expand=Parent`;
        }
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            let res = await request.get(requestOptions);
            let errorMessage = '';
            if (args.options.name) {
                if (res.value.length === 0) {
                    errorMessage = `Content type with name ${args.options.name} not found`;
                }
                else {
                    res = res.value[0];
                }
            }
            if (args.options.id && res['odata.null'] === true) {
                errorMessage = `Content type with ID ${args.options.id} not found`;
            }
            if (errorMessage) {
                throw errorMessage;
            }
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoContentTypeGetCommand_instances = new WeakSet(), _SpoContentTypeGetCommand_initTelemetry = function _SpoContentTypeGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _SpoContentTypeGetCommand_initOptions = function _SpoContentTypeGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listTitle [listTitle]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    });
}, _SpoContentTypeGetCommand_initValidators = function _SpoContentTypeGetCommand_initValidators() {
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
}, _SpoContentTypeGetCommand_initTypes = function _SpoContentTypeGetCommand_initTypes() {
    this.types.string.push('id', 'i');
}, _SpoContentTypeGetCommand_initOptionSets = function _SpoContentTypeGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new SpoContentTypeGetCommand();
//# sourceMappingURL=contenttype-get.js.map