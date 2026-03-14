var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListWebhookGetCommand_instances, _SpoListWebhookGetCommand_initTelemetry, _SpoListWebhookGetCommand_initOptions, _SpoListWebhookGetCommand_initValidators, _SpoListWebhookGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListWebhookGetCommand extends SpoCommand {
    get name() {
        return commands.LIST_WEBHOOK_GET;
    }
    get description() {
        return 'Gets information about the specific webhook';
    }
    constructor() {
        super();
        _SpoListWebhookGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListWebhookGetCommand_instances, "m", _SpoListWebhookGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookGetCommand_instances, "m", _SpoListWebhookGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookGetCommand_instances, "m", _SpoListWebhookGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookGetCommand_instances, "m", _SpoListWebhookGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            const list = (args.options.listId || args.options.listTitle || args.options.listUrl);
            await logger.logToStderr(`Retrieving information for webhook ${args.options.id} belonging to list ${list} in site at ${args.options.webUrl}...`);
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
            method: 'GET',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            if (this.verbose) {
                await logger.logToStderr('Specified webhook not found');
            }
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListWebhookGetCommand_instances = new WeakSet(), _SpoListWebhookGetCommand_initTelemetry = function _SpoListWebhookGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListWebhookGetCommand_initOptions = function _SpoListWebhookGetCommand_initOptions() {
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
    });
}, _SpoListWebhookGetCommand_initValidators = function _SpoListWebhookGetCommand_initValidators() {
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
}, _SpoListWebhookGetCommand_initOptionSets = function _SpoListWebhookGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListWebhookGetCommand();
//# sourceMappingURL=list-webhook-get.js.map