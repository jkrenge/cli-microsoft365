var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemSetCommand_instances, _SpoListItemSetCommand_initTelemetry, _SpoListItemSetCommand_initOptions, _SpoListItemSetCommand_initValidators, _SpoListItemSetCommand_initTypes, _SpoListItemSetCommand_initOptionSets;
import request from '../../../../request.js';
import { basic } from '../../../../utils/basic.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemSetCommand extends SpoCommand {
    allowUnknownOptions() {
        return true;
    }
    get name() {
        return commands.LISTITEM_SET;
    }
    get description() {
        return 'Updates a list item in the specified list';
    }
    constructor() {
        super();
        _SpoListItemSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemSetCommand_instances, "m", _SpoListItemSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemSetCommand_instances, "m", _SpoListItemSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemSetCommand_instances, "m", _SpoListItemSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemSetCommand_instances, "m", _SpoListItemSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoListItemSetCommand_instances, "m", _SpoListItemSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let contentTypeName = '';
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
            if (args.options.contentType) {
                if (this.verbose) {
                    await logger.logToStderr(`Getting content types for list...`);
                }
                const requestOptions = {
                    url: `${requestUrl}/contenttypes?$select=Name,Id`,
                    headers: {
                        'accept': 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const contentTypes = await request.get(requestOptions);
                if (this.debug) {
                    await logger.logToStderr('content type lookup response...');
                    await logger.logToStderr(contentTypes);
                }
                const foundContentType = await basic.asyncFilter(contentTypes.value, async (ct) => {
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
            const properties = this.mapRequestBody(args.options);
            const item = args.options.systemUpdate ?
                await spo.systemUpdateListItem(requestUrl, +args.options.id, logger, this.verbose, properties, contentTypeName)
                : await spo.updateListItem(requestUrl, args.options.id, properties, contentTypeName);
            delete item.ID;
            await logger.log(item);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    mapRequestBody(options) {
        const filteredData = {};
        const excludeOptions = [
            'listTitle',
            'listId',
            'listUrl',
            'webUrl',
            'id',
            'contentType',
            'systemUpdate',
            'debug',
            'verbose',
            'output',
            's',
            'i',
            'o',
            'u',
            't',
            '_'
        ];
        for (const key of Object.keys(options)) {
            if (!excludeOptions.includes(key)) {
                filteredData[key] = options[key];
            }
        }
        return filteredData;
    }
}
_SpoListItemSetCommand_instances = new WeakSet(), _SpoListItemSetCommand_initTelemetry = function _SpoListItemSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            contentType: typeof args.options.contentType !== 'undefined',
            systemUpdate: typeof args.options.systemUpdate !== 'undefined'
        });
    });
}, _SpoListItemSetCommand_initOptions = function _SpoListItemSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-c, --contentType [contentType]'
    }, {
        option: '-s, --systemUpdate'
    });
}, _SpoListItemSetCommand_initValidators = function _SpoListItemSetCommand_initValidators() {
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
}, _SpoListItemSetCommand_initTypes = function _SpoListItemSetCommand_initTypes() {
    this.types.string.push('webUrl', 'listId', 'listTitle', 'listUrl', 'id', 'contentType');
    this.types.boolean.push('systemUpdate');
}, _SpoListItemSetCommand_initOptionSets = function _SpoListItemSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemSetCommand();
//# sourceMappingURL=listitem-set.js.map