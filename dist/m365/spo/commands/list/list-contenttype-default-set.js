var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListContentTypeDefaultSetCommand_instances, _SpoListContentTypeDefaultSetCommand_initTelemetry, _SpoListContentTypeDefaultSetCommand_initOptions, _SpoListContentTypeDefaultSetCommand_initValidators, _SpoListContentTypeDefaultSetCommand_initTypes, _SpoListContentTypeDefaultSetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListContentTypeDefaultSetCommand extends SpoCommand {
    get name() {
        return commands.LIST_CONTENTTYPE_DEFAULT_SET;
    }
    get description() {
        return 'Sets the default content type for a list';
    }
    constructor() {
        super();
        _SpoListContentTypeDefaultSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListContentTypeDefaultSetCommand_instances, "m", _SpoListContentTypeDefaultSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeDefaultSetCommand_instances, "m", _SpoListContentTypeDefaultSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeDefaultSetCommand_instances, "m", _SpoListContentTypeDefaultSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeDefaultSetCommand_instances, "m", _SpoListContentTypeDefaultSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeDefaultSetCommand_instances, "m", _SpoListContentTypeDefaultSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let baseUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.listId) {
            baseUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            baseUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            baseUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        if (this.verbose) {
            await logger.logToStderr('Retrieving content type order...');
        }
        try {
            const contentTypeOrder = await this.getContentTypeOrder(baseUrl, logger);
            // see if the specified content type is among the registered content types
            // if it is, it means it's visible
            const contentTypeIndex = contentTypeOrder.findIndex(ct => ct.StringValue.toUpperCase() === args.options.contentTypeId.toUpperCase());
            if (contentTypeIndex > -1) {
                if (this.debug) {
                    await logger.logToStderr(`Content type ${args.options.contentTypeId} is visible in the list`);
                }
                // content type is in the list and is visible in the menu
                if (contentTypeIndex === 0) {
                    if (this.verbose) {
                        await logger.logToStderr(`Content type ${args.options.contentTypeId} is already set as default`);
                    }
                }
                else {
                    if (this.verbose) {
                        await logger.logToStderr(`Setting content type ${args.options.contentTypeId} as default...`);
                    }
                    // remove content type from the order array so that we can put it at
                    // the beginning to make it default content type          
                    contentTypeOrder.splice(contentTypeIndex, 1);
                    contentTypeOrder.unshift({
                        StringValue: args.options.contentTypeId
                    });
                    await this.updateContentTypeOrder(baseUrl, contentTypeOrder);
                }
            }
            else {
                if (this.debug) {
                    await logger.logToStderr(`Content type ${args.options.contentTypeId} is not visible in the list`);
                }
                if (this.verbose) {
                    await logger.logToStderr('Retrieving list content types...');
                }
                const contentTypes = await this.getListContentTypes(baseUrl);
                if (!contentTypes.find(ct => ct.toUpperCase() === args.options.contentTypeId.toUpperCase())) {
                    throw `Content type ${args.options.contentTypeId} missing in the list. Add the content type to the list first and try again.`;
                }
                if (this.verbose) {
                    await logger.logToStderr(`Setting content type ${args.options.contentTypeId} as default...`);
                }
                contentTypeOrder.unshift({
                    StringValue: args.options.contentTypeId
                });
                await this.updateContentTypeOrder(baseUrl, contentTypeOrder);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getContentTypeOrder(baseUrl, logger) {
        const requestOptions = {
            url: `${baseUrl}/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        let uniqueContentTypeOrder = response.ContentTypeOrder;
        if (response.UniqueContentTypeOrder !== null) {
            if (this.debug) {
                await logger.logToStderr('Using unique content type order');
            }
            uniqueContentTypeOrder = response.UniqueContentTypeOrder;
        }
        else {
            if (this.debug) {
                await logger.logToStderr('Unique content type order not defined. Using content type order');
            }
        }
        return uniqueContentTypeOrder;
    }
    async updateContentTypeOrder(baseUrl, contentTypeOrder) {
        const requestOptions = {
            url: `${baseUrl}/RootFolder`,
            headers: {
                'accept': 'application/json;odata=nometadata',
                'content-type': 'application/json;odata=nometadata',
                'x-http-method': 'MERGE'
            },
            data: {
                UniqueContentTypeOrder: contentTypeOrder
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async getListContentTypes(baseUrl) {
        const requestOptions = {
            url: `${baseUrl}/ContentTypes?$select=Id`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return response.value.map(ct => ct.Id.StringValue);
    }
}
_SpoListContentTypeDefaultSetCommand_instances = new WeakSet(), _SpoListContentTypeDefaultSetCommand_initTelemetry = function _SpoListContentTypeDefaultSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListContentTypeDefaultSetCommand_initOptions = function _SpoListContentTypeDefaultSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-c, --contentTypeId <contentTypeId>'
    });
}, _SpoListContentTypeDefaultSetCommand_initValidators = function _SpoListContentTypeDefaultSetCommand_initValidators() {
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
}, _SpoListContentTypeDefaultSetCommand_initTypes = function _SpoListContentTypeDefaultSetCommand_initTypes() {
    this.types.string.push('contentTypeId', 'c');
}, _SpoListContentTypeDefaultSetCommand_initOptionSets = function _SpoListContentTypeDefaultSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListContentTypeDefaultSetCommand();
//# sourceMappingURL=list-contenttype-default-set.js.map