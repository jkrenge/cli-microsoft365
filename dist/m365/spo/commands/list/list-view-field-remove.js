var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListViewFieldRemoveCommand_instances, _SpoListViewFieldRemoveCommand_initTelemetry, _SpoListViewFieldRemoveCommand_initOptions, _SpoListViewFieldRemoveCommand_initValidators, _SpoListViewFieldRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListViewFieldRemoveCommand extends SpoCommand {
    get name() {
        return commands.LIST_VIEW_FIELD_REMOVE;
    }
    get description() {
        return 'Removes the specified field from list view';
    }
    constructor() {
        super();
        _SpoListViewFieldRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListViewFieldRemoveCommand_instances, "m", _SpoListViewFieldRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListViewFieldRemoveCommand_instances, "m", _SpoListViewFieldRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListViewFieldRemoveCommand_instances, "m", _SpoListViewFieldRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListViewFieldRemoveCommand_instances, "m", _SpoListViewFieldRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeFieldFromView = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Getting field ${args.options.id || args.options.title}...`);
                }
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
                const field = await this.getField(args.options, listSelector);
                if (this.verbose) {
                    await logger.logToStderr(`Removing field ${args.options.id || args.options.title} from view ${args.options.viewId || args.options.viewTitle}...`);
                }
                const viewSelector = args.options.viewId ? `('${formatting.encodeQueryParameter(args.options.viewId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle)}')`;
                const postRequestUrl = `${args.options.webUrl}/_api/web/${listSelector}/views${viewSelector}/viewfields/removeviewfield('${field.InternalName}')`;
                const postRequestOptions = {
                    url: postRequestUrl,
                    headers: {
                        'accept': 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                await request.post(postRequestOptions);
                // REST post call doesn't return anything
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeFieldFromView();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the field ${args.options.id || args.options.title} from the view ${args.options.viewId || args.options.viewTitle} from list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}?` });
            if (result) {
                await removeFieldFromView();
            }
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
_SpoListViewFieldRemoveCommand_instances = new WeakSet(), _SpoListViewFieldRemoveCommand_initTelemetry = function _SpoListViewFieldRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            viewId: typeof args.options.viewId !== 'undefined',
            viewTitle: typeof args.options.viewTitle !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoListViewFieldRemoveCommand_initOptions = function _SpoListViewFieldRemoveCommand_initOptions() {
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
        option: '-f, --force'
    });
}, _SpoListViewFieldRemoveCommand_initValidators = function _SpoListViewFieldRemoveCommand_initValidators() {
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
                return `${args.options.viewId} is not a valid GUID`;
            }
        }
        return true;
    });
}, _SpoListViewFieldRemoveCommand_initOptionSets = function _SpoListViewFieldRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['viewId', 'viewTitle'] }, { options: ['id', 'title'] });
};
export default new SpoListViewFieldRemoveCommand();
//# sourceMappingURL=list-view-field-remove.js.map