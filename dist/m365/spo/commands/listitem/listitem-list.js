var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemListCommand_instances, _SpoListItemListCommand_initTelemetry, _SpoListItemListCommand_initOptions, _SpoListItemListCommand_initValidators, _SpoListItemListCommand_initOptionSets, _SpoListItemListCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemListCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_LIST;
    }
    get description() {
        return 'Gets a list of items from the specified list';
    }
    constructor() {
        super();
        _SpoListItemListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemListCommand_instances, "m", _SpoListItemListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemListCommand_instances, "m", _SpoListItemListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemListCommand_instances, "m", _SpoListItemListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemListCommand_instances, "m", _SpoListItemListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoListItemListCommand_instances, "m", _SpoListItemListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        let listApiUrl = `${args.options.webUrl}/_api/web`;
        if (args.options.listId) {
            listApiUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            listApiUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            listApiUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        try {
            const listItems = args.options.camlQuery
                ? await this.getItemsUsingCAMLQuery(logger, args.options, listApiUrl)
                : await this.getItems(logger, args.options, listApiUrl);
            listItems.forEach(v => delete v['ID']);
            await logger.log(listItems);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getItems(logger, options, listApiUrl) {
        if (this.verbose) {
            await logger.logToStderr(`Getting list items`);
        }
        const queryParams = [];
        const fieldsArray = options.fields ? options.fields.split(",")
            : (!options.output || cli.shouldTrimOutput(options.output)) ? ["Title", "Id"] : [];
        const expandFieldsArray = this.getExpandFieldsArray(fieldsArray);
        const skipTokenId = await this.getLastItemIdForPage(logger, options, listApiUrl);
        queryParams.push(`$top=${options.pageSize || 5000}`);
        if (options.filter) {
            queryParams.push(`$filter=${encodeURIComponent(options.filter)}`);
        }
        if (expandFieldsArray.length > 0) {
            queryParams.push(`$expand=${expandFieldsArray.join(",")}`);
        }
        if (fieldsArray.length > 0) {
            queryParams.push(`$select=${formatting.encodeQueryParameter(fieldsArray.join(","))}`);
        }
        if (skipTokenId !== undefined) {
            queryParams.push(`$skiptoken=Paged=TRUE%26p_ID=${skipTokenId}`);
        }
        // If skiptoken is not found, then we are past the last page
        if (options.pageNumber && Number(options.pageNumber) > 0 && skipTokenId === undefined) {
            return [];
        }
        if (!options.pageSize) {
            return await odata.getAllItems(`${listApiUrl}/items?${queryParams.join('&')}`);
        }
        else {
            const requestOptions = {
                url: `${listApiUrl}/items?${queryParams.join('&')}`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const listItemCollection = await request.get(requestOptions);
            return listItemCollection.value;
        }
    }
    async getItemsUsingCAMLQuery(logger, options, listApiUrl) {
        const formDigestValue = (await spo.getRequestDigest(options.webUrl)).FormDigestValue;
        if (this.verbose) {
            await logger.logToStderr(`Getting list items using CAML query`);
        }
        const items = [];
        let skipTokenId = undefined;
        do {
            const requestBody = {
                "query": {
                    "ViewXml": options.camlQuery,
                    "AllowIncrementalResults": true
                }
            };
            if (skipTokenId !== undefined) {
                requestBody.query.ListItemCollectionPosition = {
                    "PagingInfo": `Paged=TRUE&p_ID=${skipTokenId}`
                };
            }
            const requestOptions = {
                url: `${listApiUrl}/GetItems`,
                headers: {
                    'accept': 'application/json;odata=nometadata',
                    'X-RequestDigest': formDigestValue
                },
                responseType: 'json',
                data: requestBody
            };
            const listItemInstances = await request.post(requestOptions);
            skipTokenId = listItemInstances.value.length > 0 ? listItemInstances.value[listItemInstances.value.length - 1].Id : undefined;
            items.push(...listItemInstances.value);
        } while (skipTokenId !== undefined);
        return items;
    }
    getExpandFieldsArray(fieldsArray) {
        const fieldsWithSlash = fieldsArray.filter(item => item.includes('/'));
        const fieldsToExpand = fieldsWithSlash.map(e => e.split('/')[0]);
        const expandFieldsArray = fieldsToExpand.filter((item, pos) => fieldsToExpand.indexOf(item) === pos);
        return expandFieldsArray;
    }
    async getLastItemIdForPage(logger, options, listApiUrl) {
        if (!options.pageNumber || Number(options.pageNumber) === 0) {
            return undefined;
        }
        if (this.verbose) {
            await logger.logToStderr(`Getting skipToken Id for page ${options.pageNumber}`);
        }
        const rowLimit = `$top=${Number(options.pageSize) * Number(options.pageNumber)}`;
        const filter = options.filter ? `$filter=${encodeURIComponent(options.filter)}` : ``;
        const requestOptions = {
            url: `${listApiUrl}/items?$select=Id&${rowLimit}&${filter}`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response.value[response.value.length - 1]?.Id;
    }
}
_SpoListItemListCommand_instances = new WeakSet(), _SpoListItemListCommand_initTelemetry = function _SpoListItemListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            fields: typeof args.options.fields !== 'undefined',
            filter: typeof args.options.filter !== 'undefined',
            pageNumber: typeof args.options.pageNumber !== 'undefined',
            pageSize: typeof args.options.pageSize !== 'undefined',
            camlQuery: typeof args.options.camlQuery !== 'undefined'
        });
    });
}, _SpoListItemListCommand_initOptions = function _SpoListItemListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-s, --pageSize [pageSize]'
    }, {
        option: '-n, --pageNumber [pageNumber]'
    }, {
        option: '-q, --camlQuery [camlQuery]'
    }, {
        option: '--fields [fields]'
    }, {
        option: '-l, --filter [filter]'
    });
}, _SpoListItemListCommand_initValidators = function _SpoListItemListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.camlQuery && args.options.fields) {
            return `You cannot use the fields-option when specifying a camlQuery`;
        }
        if (args.options.camlQuery && args.options.pageSize) {
            return `You cannot use the pageSize-option when specifying a camlQuery`;
        }
        if (args.options.camlQuery && args.options.pageNumber) {
            return `You cannot use the pageNumber-option when specifying a camlQuery`;
        }
        if (args.options.pageSize && isNaN(Number(args.options.pageSize))) {
            return `pageSize ${args.options.pageSize} must be numeric`;
        }
        if (args.options.pageNumber && !args.options.pageSize) {
            return `pageSize must be specified if pageNumber is specified`;
        }
        if (args.options.pageNumber && isNaN(Number(args.options.pageNumber))) {
            return `pageNumber must be numeric`;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoListItemListCommand_initOptionSets = function _SpoListItemListCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
}, _SpoListItemListCommand_initTypes = function _SpoListItemListCommand_initTypes() {
    this.types.string.push('webUrl', 'camlQuery', 'pageSize', 'pageNumber', 'fields', 'filter');
};
export default new SpoListItemListCommand();
//# sourceMappingURL=listitem-list.js.map