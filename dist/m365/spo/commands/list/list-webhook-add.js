var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListWebhookAddCommand_instances, _SpoListWebhookAddCommand_initTelemetry, _SpoListWebhookAddCommand_initOptions, _SpoListWebhookAddCommand_initValidators, _SpoListWebhookAddCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
const expirationDateTimeMaxDays = 180;
const maxExpirationDateTime = new Date();
// 180 days from now is the maximum expiration date for a webhook
maxExpirationDateTime.setDate(maxExpirationDateTime.getDate() + expirationDateTimeMaxDays);
class SpoListWebhookAddCommand extends SpoCommand {
    get name() {
        return commands.LIST_WEBHOOK_ADD;
    }
    get description() {
        return 'Adds a new webhook to the specified list';
    }
    constructor() {
        super();
        _SpoListWebhookAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListWebhookAddCommand_instances, "m", _SpoListWebhookAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookAddCommand_instances, "m", _SpoListWebhookAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookAddCommand_instances, "m", _SpoListWebhookAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookAddCommand_instances, "m", _SpoListWebhookAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding webhook to list ${args.options.listId || args.options.listTitle || args.options.listUrl} located at site ${args.options.webUrl}...`);
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
        const requestBody = {};
        requestBody.resource = args.options.listId || args.options.listTitle || args.options.listUrl;
        requestBody.notificationUrl = args.options.notificationUrl;
        // If no expiration date has been provided we will default to the
        // maximum expiration date of 180 days from now 
        requestBody.expirationDateTime = args.options.expirationDateTime
            ? new Date(args.options.expirationDateTime).toISOString()
            : maxExpirationDateTime.toISOString();
        if (args.options.clientState) {
            requestBody.clientState = args.options.clientState;
        }
        const requestOptions = {
            url: requestUrl,
            method: 'POST',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: requestBody,
            responseType: 'json'
        };
        try {
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListWebhookAddCommand_instances = new WeakSet(), _SpoListWebhookAddCommand_initTelemetry = function _SpoListWebhookAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            expirationDateTime: typeof args.options.expirationDateTime !== 'undefined',
            clientState: typeof args.options.clientState !== 'undefined'
        });
    });
}, _SpoListWebhookAddCommand_initOptions = function _SpoListWebhookAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-n, --notificationUrl <notificationUrl>'
    }, {
        option: '-e, --expirationDateTime [expirationDateTime]'
    }, {
        option: '-c, --clientState [clientState]'
    });
}, _SpoListWebhookAddCommand_initValidators = function _SpoListWebhookAddCommand_initValidators() {
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
        const parsedDateTime = Date.parse(args.options.expirationDateTime);
        if (args.options.expirationDateTime && !(!parsedDateTime) !== true) {
            return `Provide the date in one of the following formats:
          'YYYY-MM-DD'
          'YYYY-MM-DDThh:mm'
          'YYYY-MM-DDThh:mmZ'
          'YYYY-MM-DDThh:mm±hh:mm'`;
        }
        if (parsedDateTime < Date.now() || new Date(parsedDateTime) >= maxExpirationDateTime) {
            return `Provide an expiration date which is a date time in the future and within 6 months from now`;
        }
        return true;
    });
}, _SpoListWebhookAddCommand_initOptionSets = function _SpoListWebhookAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListWebhookAddCommand();
//# sourceMappingURL=list-webhook-add.js.map