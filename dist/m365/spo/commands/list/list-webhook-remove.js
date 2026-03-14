var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListWebhookRemoveCommand_instances, _SpoListWebhookRemoveCommand_initTelemetry, _SpoListWebhookRemoveCommand_initOptions, _SpoListWebhookRemoveCommand_initValidators, _SpoListWebhookRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListWebhookRemoveCommand extends SpoCommand {
    get name() {
        return commands.LIST_WEBHOOK_REMOVE;
    }
    get description() {
        return 'Removes the specified webhook from the list';
    }
    constructor() {
        super();
        _SpoListWebhookRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListWebhookRemoveCommand_instances, "m", _SpoListWebhookRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookRemoveCommand_instances, "m", _SpoListWebhookRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookRemoveCommand_instances, "m", _SpoListWebhookRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookRemoveCommand_instances, "m", _SpoListWebhookRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeWebhook = async () => {
            if (this.verbose) {
                const list = (args.options.listId || args.options.listId || args.options.listUrl);
                await logger.logToStderr(`Webhook ${args.options.id} is about to be removed from list ${list} located at site ${args.options.webUrl}...`);
            }
            let requestUrl = `${args.options.webUrl}/_api/web`;
            if (args.options.listId) {
                requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
            }
            else if (args.options.listTitle) {
                requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
            }
            else if (args.options.listUrl) {
                const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
                requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
            }
            const requestOptions = {
                url: requestUrl,
                method: 'DELETE',
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            try {
                await request.delete(requestOptions);
                // REST delete call doesn't return anything
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeWebhook();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove webhook ${args.options.id} from list ${args.options.listTitle || args.options.listId || args.options.listUrl} located at site ${args.options.webUrl}?` });
            if (result) {
                await removeWebhook();
            }
        }
    }
}
_SpoListWebhookRemoveCommand_instances = new WeakSet(), _SpoListWebhookRemoveCommand_initTelemetry = function _SpoListWebhookRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoListWebhookRemoveCommand_initOptions = function _SpoListWebhookRemoveCommand_initOptions() {
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
}, _SpoListWebhookRemoveCommand_initValidators = function _SpoListWebhookRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
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
}, _SpoListWebhookRemoveCommand_initOptionSets = function _SpoListWebhookRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListWebhookRemoveCommand();
//# sourceMappingURL=list-webhook-remove.js.map