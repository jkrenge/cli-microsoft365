var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemAddCommand_instances, _SpoListItemAddCommand_initTelemetry, _SpoListItemAddCommand_initOptions, _SpoListItemAddCommand_initValidators, _SpoListItemAddCommand_initTypes, _SpoListItemAddCommand_initOptionSets;
import os from 'os';
import request from '../../../../request.js';
import { basic } from '../../../../utils/basic.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemAddCommand extends SpoCommand {
    allowUnknownOptions() {
        return true;
    }
    get name() {
        return commands.LISTITEM_ADD;
    }
    get description() {
        return 'Creates a list item in the specified list';
    }
    constructor() {
        super();
        _SpoListItemAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemAddCommand_instances, "m", _SpoListItemAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemAddCommand_instances, "m", _SpoListItemAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemAddCommand_instances, "m", _SpoListItemAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemAddCommand_instances, "m", _SpoListItemAddCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoListItemAddCommand_instances, "m", _SpoListItemAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
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
            let contentTypeName = '';
            let targetFolderServerRelativeUrl = '';
            if (this.verbose) {
                await logger.logToStderr(`Getting content types for list ${args.options.listId || args.options.listTitle || args.options.listUrl}...`);
            }
            let requestOptions = {
                url: `${requestUrl}/contenttypes?$select=Name,Id`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const ctypes = await request.get(requestOptions);
            if (args.options.contentType) {
                const foundContentType = await basic.asyncFilter(ctypes.value, async (ct) => {
                    const contentTypeMatch = ct.Id.StringValue === args.options.contentType || ct.Name === args.options.contentType;
                    if (this.debug) {
                        await logger.logToStderr(`Checking content type value [${ct.Name}]: ${contentTypeMatch}`);
                    }
                    return contentTypeMatch;
                });
                if (this.debug) {
                    await logger.logToStderr('content type filter output...');
                    await logger.logToStderr(foundContentType);
                }
                if (foundContentType.length > 0) {
                    contentTypeName = foundContentType[0].Name;
                }
                // After checking for content types, throw an error if the name is blank
                if (!contentTypeName || contentTypeName === '') {
                    throw `Specified content type '${args.options.contentType}' doesn't exist on the target list`;
                }
                if (this.debug) {
                    await logger.logToStderr(`using content type name: ${contentTypeName}`);
                }
            }
            if (args.options.folder) {
                if (this.debug) {
                    await logger.logToStderr('setting up folder lookup response ...');
                }
                requestOptions = {
                    url: `${requestUrl}/rootFolder`,
                    headers: {
                        'accept': 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const rootFolderResponse = await request.get(requestOptions);
                targetFolderServerRelativeUrl = urlUtil.getServerRelativePath(rootFolderResponse["ServerRelativeUrl"], args.options.folder);
                await spo.ensureFolder(args.options.webUrl, targetFolderServerRelativeUrl, logger, this.debug);
            }
            if (this.verbose) {
                await logger.logToStderr(`Creating item in list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site ${args.options.webUrl}...`);
            }
            const requestBody = {
                formValues: this.mapRequestBody(args.options)
            };
            if (args.options.folder) {
                requestBody.listItemCreateInfo = {
                    FolderPath: {
                        DecodedUrl: targetFolderServerRelativeUrl
                    }
                };
            }
            if (args.options.contentType && contentTypeName !== '') {
                if (this.debug) {
                    await logger.logToStderr(`Specifying content type name [${contentTypeName}] in request body`);
                }
                requestBody.formValues.push({
                    FieldName: 'ContentType',
                    FieldValue: contentTypeName
                });
            }
            requestOptions = {
                url: `${requestUrl}/AddValidateUpdateItemUsingPath()`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: requestBody,
                responseType: 'json'
            };
            const response = await request.post(requestOptions);
            // Response is from /AddValidateUpdateItemUsingPath POST call, perform get on added item to get all field values
            const fieldValues = response.value;
            if (fieldValues.some(f => f.HasException)) {
                throw `Creating the item failed with the following errors: ${os.EOL}${fieldValues.filter(f => f.HasException).map(f => { return `- ${f.FieldName} - ${f.ErrorMessage}`; }).join(os.EOL)}`;
            }
            const idField = fieldValues.filter((thisField) => {
                return (thisField.FieldName === "Id");
            });
            if (this.debug) {
                await logger.logToStderr(`Field values returned:`);
                await logger.logToStderr(fieldValues);
                await logger.logToStderr(`Id returned by AddValidateUpdateItemUsingPath: ${idField[0].FieldValue}`);
            }
            requestOptions = {
                url: `${requestUrl}/items(${idField[0].FieldValue})`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const item = await request.get(requestOptions);
            delete item.ID;
            await logger.log(item);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    mapRequestBody(options) {
        const requestBody = [];
        const excludeOptions = [
            'listTitle',
            'listId',
            'listUrl',
            'webUrl',
            'contentType',
            'folder',
            'debug',
            'verbose',
            'output',
            '_'
        ];
        Object.keys(options).forEach(key => {
            if (excludeOptions.indexOf(key) === -1) {
                requestBody.push({ FieldName: key, FieldValue: `${options[key]}` });
            }
        });
        return requestBody;
    }
}
_SpoListItemAddCommand_instances = new WeakSet(), _SpoListItemAddCommand_initTelemetry = function _SpoListItemAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            contentType: typeof args.options.contentType !== 'undefined',
            folder: typeof args.options.folder !== 'undefined'
        });
    });
}, _SpoListItemAddCommand_initOptions = function _SpoListItemAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-c, --contentType [contentType]'
    }, {
        option: '--folder [folder]'
    });
}, _SpoListItemAddCommand_initValidators = function _SpoListItemAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        return true;
    });
}, _SpoListItemAddCommand_initTypes = function _SpoListItemAddCommand_initTypes() {
    this.types.string.push('webUrl', 'listId', 'listTitle', 'listUrl', 'contentType', 'folder');
}, _SpoListItemAddCommand_initOptionSets = function _SpoListItemAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemAddCommand();
//# sourceMappingURL=listitem-add.js.map