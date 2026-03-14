var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListViewFieldSetCommand_instances, _SpoListViewFieldSetCommand_initTelemetry, _SpoListViewFieldSetCommand_initOptions, _SpoListViewFieldSetCommand_initValidators, _SpoListViewFieldSetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListViewFieldSetCommand extends SpoCommand {
    get name() {
        return commands.LIST_VIEW_FIELD_SET;
    }
    get description() {
        return 'Updates existing column in an existing view (eg. move to a specific position).';
    }
    constructor() {
        super();
        _SpoListViewFieldSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListViewFieldSetCommand_instances, "m", _SpoListViewFieldSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListViewFieldSetCommand_instances, "m", _SpoListViewFieldSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListViewFieldSetCommand_instances, "m", _SpoListViewFieldSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListViewFieldSetCommand_instances, "m", _SpoListViewFieldSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let listSelector = '';
        if (args.options.listId) {
            listSelector = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            listSelector = `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            listSelector = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        const viewSelector = args.options.viewId ? `('${formatting.encodeQueryParameter(args.options.viewId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle)}')`;
        if (this.verbose) {
            await logger.logToStderr(`Getting field ${args.options.id || args.options.title}...`);
        }
        try {
            const field = await this.getField(args.options, listSelector);
            if (this.verbose) {
                await logger.logToStderr(`Moving the field ${args.options.id || args.options.title} in view ${args.options.viewId || args.options.viewTitle} to position ${args.options.position}...`);
            }
            const moveRequestUrl = `${args.options.webUrl}/_api/web/${listSelector}/views${viewSelector}/viewfields/moveviewfieldto`;
            const moveRequestOptions = {
                url: moveRequestUrl,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: {
                    field: field.InternalName,
                    index: args.options.position
                },
                responseType: 'json'
            };
            await request.post(moveRequestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getField(options, listSelector) {
        const fieldSelector = options.id ? `/getbyid('${formatting.encodeQueryParameter(options.id)}')` : `/getbyinternalnameortitle('${formatting.encodeQueryParameter(options.title)}')`;
        const getRequestUrl = `${options.webUrl}/_api/web/${listSelector}/fields${fieldSelector}`;
        const requestOptions = {
            url: getRequestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
}
_SpoListViewFieldSetCommand_instances = new WeakSet(), _SpoListViewFieldSetCommand_initTelemetry = function _SpoListViewFieldSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            viewId: typeof args.options.viewId !== 'undefined',
            viewTitle: typeof args.options.viewTitle !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined'
        });
    });
}, _SpoListViewFieldSetCommand_initOptions = function _SpoListViewFieldSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '--viewId [viewId]'
    }, {
        option: '--viewTitle [viewTitle]'
    }, {
        option: '--id [id]'
    }, {
        option: '--title [title]'
    }, {
        option: '--position <position>'
    });
}, _SpoListViewFieldSetCommand_initValidators = function _SpoListViewFieldSetCommand_initValidators() {
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
        if (args.options.viewId) {
            if (!validation.isValidGuid(args.options.viewId)) {
                return `${args.options.viewId} is not a valid GUID`;
            }
        }
        if (args.options.id) {
            if (!validation.isValidGuid(args.options.id)) {
                return `${args.options.id} is not a valid GUID`;
            }
        }
        const position = parseInt(args.options.position);
        if (isNaN(position)) {
            return `${args.options.position} is not a number`;
        }
        return true;
    });
}, _SpoListViewFieldSetCommand_initOptionSets = function _SpoListViewFieldSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['viewId', 'viewTitle'] }, { options: ['id', 'title'] });
};
export default new SpoListViewFieldSetCommand();
//# sourceMappingURL=list-view-field-set.js.map