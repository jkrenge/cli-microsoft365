var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoContentTypeSyncCommand_instances, _SpoContentTypeSyncCommand_initTelemetry, _SpoContentTypeSyncCommand_initOptions, _SpoContentTypeSyncCommand_initValidators, _SpoContentTypeSyncCommand_initTypes, _SpoContentTypeSyncCommand_initOptionSets;
import { formatting } from '../../../../utils/formatting.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
class SpoContentTypeSyncCommand extends SpoCommand {
    get name() {
        return commands.CONTENTTYPE_SYNC;
    }
    get description() {
        return 'Adds a published content type from the content type hub to a site or syncs its latest changes';
    }
    constructor() {
        super();
        _SpoContentTypeSyncCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoContentTypeSyncCommand_instances, "m", _SpoContentTypeSyncCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeSyncCommand_instances, "m", _SpoContentTypeSyncCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeSyncCommand_instances, "m", _SpoContentTypeSyncCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeSyncCommand_instances, "m", _SpoContentTypeSyncCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoContentTypeSyncCommand_instances, "m", _SpoContentTypeSyncCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const { listId, listTitle, listUrl, webUrl } = args.options;
        const url = new URL(webUrl);
        const baseUrl = 'https://graph.microsoft.com/v1.0/sites/';
        try {
            const siteUrl = url.pathname === '/' ? url.host : await spo.getSiteIdByMSGraph(webUrl, logger, this.verbose);
            const listPath = listId || listTitle || listUrl ? `/lists/${listId || listTitle || await this.getListIdByUrl(webUrl, listUrl, logger)}` : '';
            const contentTypeId = await this.getContentTypeId(baseUrl, url, args.options, logger);
            if (this.verbose) {
                await logger.logToStderr(`Adding or syncing the content type...`);
            }
            const requestOptions = {
                url: `${baseUrl}${siteUrl}${listPath}/contenttypes/addCopyFromContentTypeHub`,
                headers: {
                    'accept': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false'
                },
                responseType: 'json',
                data: {
                    contentTypeId: contentTypeId
                }
            };
            const response = await request.post(requestOptions);
            // The endpoint only returns a response if the content type has been added for the first time
            // When syncing, the response will be an empty string, which should not be logged.
            if (typeof response === 'object') {
                await logger.log(response);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getContentTypeId(baseUrl, url, options, logger) {
        if (options.id) {
            return options.id;
        }
        const siteId = await spo.getSiteIdByMSGraph(`${url.origin}/sites/contenttypehub`, logger, this.verbose);
        if (this.verbose) {
            await logger.logToStderr(`Retrieving content type Id by name...`);
        }
        const contentTypes = await odata.getAllItems(`${baseUrl}${siteId}/contenttypes?$filter=name eq '${options.name}'&$select=id,name`);
        if (contentTypes.length === 0) {
            throw `Content type with name ${options.name} not found.`;
        }
        return contentTypes[0].id;
    }
    async getListIdByUrl(webUrl, listUrl, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list id to sync the content type to...`);
        }
        const listServerRelativeUrl = urlUtil.getServerRelativePath(webUrl, listUrl);
        const requestOptions = {
            url: `${webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')?$select=id`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response.Id;
    }
}
_SpoContentTypeSyncCommand_instances = new WeakSet(), _SpoContentTypeSyncCommand_initTelemetry = function _SpoContentTypeSyncCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoContentTypeSyncCommand_initOptions = function _SpoContentTypeSyncCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoContentTypeSyncCommand_initValidators = function _SpoContentTypeSyncCommand_initValidators() {
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
}, _SpoContentTypeSyncCommand_initTypes = function _SpoContentTypeSyncCommand_initTypes() {
    this.types.string.push('webUrl', 'id', 'name', 'listTitle', 'listId', 'listUrl');
}, _SpoContentTypeSyncCommand_initOptionSets = function _SpoContentTypeSyncCommand_initOptionSets() {
    this.optionSets.push({
        options: ['id', 'name']
    }, {
        options: ['listId', 'listTitle', 'listUrl'],
        runsWhen: (args) => args.options.listId || args.options.listTitle || args.options.listUrl
    });
};
export default new SpoContentTypeSyncCommand();
//# sourceMappingURL=contenttype-sync.js.map