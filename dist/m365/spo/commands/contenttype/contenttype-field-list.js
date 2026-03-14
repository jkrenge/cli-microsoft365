var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoContentTypeFieldListCommand_instances, _SpoContentTypeFieldListCommand_initTelemetry, _SpoContentTypeFieldListCommand_initOptions, _SpoContentTypeFieldListCommand_initValidators, _SpoContentTypeFieldListCommand_initTypes, _SpoContentTypeFieldListCommand_initOptionSets;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoContentTypeFieldListCommand extends SpoCommand {
    get name() {
        return commands.CONTENTTYPE_FIELD_LIST;
    }
    get description() {
        return 'Lists fields for a given site or list content type';
    }
    defaultProperties() {
        return ['Id', 'Title', 'InternalName', 'Hidden'];
    }
    constructor() {
        super();
        _SpoContentTypeFieldListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldListCommand_instances, "m", _SpoContentTypeFieldListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldListCommand_instances, "m", _SpoContentTypeFieldListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldListCommand_instances, "m", _SpoContentTypeFieldListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldListCommand_instances, "m", _SpoContentTypeFieldListCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeFieldListCommand_instances, "m", _SpoContentTypeFieldListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving fields for content type '${args.options.contentTypeId || args.options.contentTypeName}' in site ${args.options.webUrl}...`);
            }
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
            requestUrl += '/contentTypes';
            const contentTypeId = await this.getContentTypeId(requestUrl, logger, args.options.contentTypeId, args.options.contentTypeName);
            requestUrl += `('${formatting.encodeQueryParameter(contentTypeId)}')/fields`;
            if (args.options.properties) {
                requestUrl += `?$select=${args.options.properties}`;
            }
            const res = await odata.getAllItems(requestUrl);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getContentTypeId(requestUrl, logger, contentTypeId, contentTypeName) {
        if (contentTypeId) {
            return contentTypeId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving content type id for content type '${contentTypeName}'.`);
        }
        const res = await odata.getAllItems(`${requestUrl}?$filter=Name eq '${formatting.encodeQueryParameter(contentTypeName)}'&$select=StringId`);
        if (res.length === 0) {
            throw `Content type with name ${contentTypeName} not found.`;
        }
        return res[0].StringId;
    }
}
_SpoContentTypeFieldListCommand_instances = new WeakSet(), _SpoContentTypeFieldListCommand_initTelemetry = function _SpoContentTypeFieldListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            contentTypeId: typeof args.options.contentTypeId !== 'undefined',
            contentTypeName: typeof args.options.contentTypeName !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            properties: typeof args.options.properties !== 'undefined'
        });
    });
}, _SpoContentTypeFieldListCommand_initOptions = function _SpoContentTypeFieldListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --contentTypeId [contentTypeId]'
    }, {
        option: '-n, --contentTypeName [contentTypeName]'
    }, {
        option: '-l, --listTitle [listTitle]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-p, --properties [properties]'
    });
}, _SpoContentTypeFieldListCommand_initValidators = function _SpoContentTypeFieldListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID for option 'listId'.`;
        }
        return true;
    });
}, _SpoContentTypeFieldListCommand_initTypes = function _SpoContentTypeFieldListCommand_initTypes() {
    this.types.string.push('webUrl', 'contentTypeId', 'contentTypeName', 'listTitle', 'listId', 'listUrl', 'properties');
}, _SpoContentTypeFieldListCommand_initOptionSets = function _SpoContentTypeFieldListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['contentTypeId', 'contentTypeName']
    }, {
        options: ['listId', 'listTitle', 'listUrl'],
        runsWhen: (args) => args.options.listId || args.options.listTitle || args.options.listUrl
    });
};
export default new SpoContentTypeFieldListCommand();
//# sourceMappingURL=contenttype-field-list.js.map