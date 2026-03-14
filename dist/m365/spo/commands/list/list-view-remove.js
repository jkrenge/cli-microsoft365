var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListViewRemoveCommand_instances, _SpoListViewRemoveCommand_initTelemetry, _SpoListViewRemoveCommand_initOptions, _SpoListViewRemoveCommand_initValidators, _SpoListViewRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListViewRemoveCommand extends SpoCommand {
    get name() {
        return commands.LIST_VIEW_REMOVE;
    }
    get description() {
        return 'Deletes the specified view from the list';
    }
    constructor() {
        super();
        _SpoListViewRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListViewRemoveCommand_instances, "m", _SpoListViewRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListViewRemoveCommand_instances, "m", _SpoListViewRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListViewRemoveCommand_instances, "m", _SpoListViewRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListViewRemoveCommand_instances, "m", _SpoListViewRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeViewFromList = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing view ${args.options.id || args.options.title} from list ${args.options.listId || args.options.listTitle || args.options.listUrl}`);
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
            const viewSelector = args.options.id ? `(guid'${formatting.encodeQueryParameter(args.options.id)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.title)}')`;
            const requestUrl = `${args.options.webUrl}/_api/web/${listSelector}/views${viewSelector}`;
            const requestOptions = {
                url: requestUrl,
                headers: {
                    'X-HTTP-Method': 'DELETE',
                    'If-Match': '*',
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            try {
                await request.post(requestOptions);
                // REST post call doesn't return anything
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeViewFromList();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the view ${args.options.id || args.options.title} from the list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site ${args.options.webUrl}?` });
            if (result) {
                await removeViewFromList();
            }
        }
    }
}
_SpoListViewRemoveCommand_instances = new WeakSet(), _SpoListViewRemoveCommand_initTelemetry = function _SpoListViewRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoListViewRemoveCommand_initOptions = function _SpoListViewRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '--id [id]'
    }, {
        option: '--title [title]'
    }, {
        option: '-f, --force'
    });
}, _SpoListViewRemoveCommand_initValidators = function _SpoListViewRemoveCommand_initValidators() {
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
        if (args.options.id) {
            if (!validation.isValidGuid(args.options.id)) {
                return `${args.options.id} is not a valid GUID`;
            }
        }
        return true;
    });
}, _SpoListViewRemoveCommand_initOptionSets = function _SpoListViewRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] }, { options: ['id', 'title'] });
};
export default new SpoListViewRemoveCommand();
//# sourceMappingURL=list-view-remove.js.map