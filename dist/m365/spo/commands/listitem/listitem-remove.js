var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRemoveCommand_instances, _SpoListItemRemoveCommand_initTelemetry, _SpoListItemRemoveCommand_initOptions, _SpoListItemRemoveCommand_initValidators, _SpoListItemRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemRemoveCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_REMOVE;
    }
    get description() {
        return 'Removes the specified list item';
    }
    constructor() {
        super();
        _SpoListItemRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRemoveCommand_instances, "m", _SpoListItemRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRemoveCommand_instances, "m", _SpoListItemRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRemoveCommand_instances, "m", _SpoListItemRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRemoveCommand_instances, "m", _SpoListItemRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeListItem = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing list item ${args.options.id} from list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
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
            requestUrl += `/items(${args.options.id})`;
            if (args.options.recycle) {
                requestUrl += `/recycle()`;
            }
            const requestOptions = {
                url: requestUrl,
                method: 'POST',
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
            await removeListItem();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the list item ${args.options.id} from list ${args.options.listId || args.options.listTitle || args.options.listUrl} located in site ${args.options.webUrl}?` });
            if (result) {
                await removeListItem();
            }
        }
    }
}
_SpoListItemRemoveCommand_instances = new WeakSet(), _SpoListItemRemoveCommand_initTelemetry = function _SpoListItemRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            recycle: !!args.options.recycle,
            force: !!args.options.force
        });
    });
}, _SpoListItemRemoveCommand_initOptions = function _SpoListItemRemoveCommand_initOptions() {
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
        option: '--recycle'
    }, {
        option: '-f, --force'
    });
}, _SpoListItemRemoveCommand_initValidators = function _SpoListItemRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const id = parseInt(args.options.id);
        if (isNaN(id)) {
            return `${args.options.id} is not a valid list item ID`;
        }
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoListItemRemoveCommand_initOptionSets = function _SpoListItemRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemRemoveCommand();
//# sourceMappingURL=listitem-remove.js.map