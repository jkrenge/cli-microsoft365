var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListContentTypeRemoveCommand_instances, _SpoListContentTypeRemoveCommand_initTelemetry, _SpoListContentTypeRemoveCommand_initOptions, _SpoListContentTypeRemoveCommand_initValidators, _SpoListContentTypeRemoveCommand_initTypes, _SpoListContentTypeRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListContentTypeRemoveCommand extends SpoCommand {
    get name() {
        return commands.LIST_CONTENTTYPE_REMOVE;
    }
    get description() {
        return 'Removes content type from list';
    }
    constructor() {
        super();
        _SpoListContentTypeRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListContentTypeRemoveCommand_instances, "m", _SpoListContentTypeRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeRemoveCommand_instances, "m", _SpoListContentTypeRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeRemoveCommand_instances, "m", _SpoListContentTypeRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeRemoveCommand_instances, "m", _SpoListContentTypeRemoveCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoListContentTypeRemoveCommand_instances, "m", _SpoListContentTypeRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeContentTypeFromList = async () => {
            if (this.verbose) {
                const list = (args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl);
                await logger.logToStderr(`Removing content type ${args.options.id} from list ${list} in site at ${args.options.webUrl}...`);
            }
            let requestUrl = `${args.options.webUrl}/_api/web/`;
            if (args.options.listId) {
                requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
            }
            else if (args.options.listTitle) {
                requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
            }
            else if (args.options.listUrl) {
                const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
                requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
            }
            const requestOptions = {
                url: `${requestUrl}/ContentTypes('${formatting.encodeQueryParameter(args.options.id)}')`,
                headers: {
                    'X-HTTP-Method': 'DELETE',
                    'If-Match': '*',
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            try {
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeContentTypeFromList();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the content type ${args.options.id} from the list ${args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl} in site ${args.options.webUrl}?` });
            if (result) {
                await removeContentTypeFromList();
            }
        }
    }
}
_SpoListContentTypeRemoveCommand_instances = new WeakSet(), _SpoListContentTypeRemoveCommand_initTelemetry = function _SpoListContentTypeRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoListContentTypeRemoveCommand_initOptions = function _SpoListContentTypeRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _SpoListContentTypeRemoveCommand_initValidators = function _SpoListContentTypeRemoveCommand_initValidators() {
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
}, _SpoListContentTypeRemoveCommand_initTypes = function _SpoListContentTypeRemoveCommand_initTypes() {
    this.types.string.push('id', 'i');
}, _SpoListContentTypeRemoveCommand_initOptionSets = function _SpoListContentTypeRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListContentTypeRemoveCommand();
//# sourceMappingURL=list-contenttype-remove.js.map