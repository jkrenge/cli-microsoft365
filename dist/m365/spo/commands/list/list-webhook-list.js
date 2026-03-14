var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListWebhookListCommand_instances, _SpoListWebhookListCommand_initTelemetry, _SpoListWebhookListCommand_initOptions, _SpoListWebhookListCommand_initValidators, _SpoListWebhookListCommand_initOptionSets;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListWebhookListCommand extends SpoCommand {
    get name() {
        return commands.LIST_WEBHOOK_LIST;
    }
    get description() {
        return 'Lists all webhooks for the specified list';
    }
    defaultProperties() {
        return ['id', 'clientState', 'expirationDateTime', 'resource'];
    }
    constructor() {
        super();
        _SpoListWebhookListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListWebhookListCommand_instances, "m", _SpoListWebhookListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookListCommand_instances, "m", _SpoListWebhookListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookListCommand_instances, "m", _SpoListWebhookListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookListCommand_instances, "m", _SpoListWebhookListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving webhook information for list ${args.options.listTitle || args.options.listId || args.options.listUrl} in site at ${args.options.webUrl}...`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web`;
        if (args.options.listId) {
            requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions`;
        }
        else if (args.options.listTitle) {
            requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/Subscriptions`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions`;
        }
        try {
            const res = await odata.getAllItems(requestUrl);
            if (res && res.length > 0) {
                res.forEach(w => {
                    w.clientState = w.clientState || '';
                });
            }
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListWebhookListCommand_instances = new WeakSet(), _SpoListWebhookListCommand_initTelemetry = function _SpoListWebhookListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListWebhookListCommand_initOptions = function _SpoListWebhookListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListWebhookListCommand_initValidators = function _SpoListWebhookListCommand_initValidators() {
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
}, _SpoListWebhookListCommand_initOptionSets = function _SpoListWebhookListCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListWebhookListCommand();
//# sourceMappingURL=list-webhook-list.js.map