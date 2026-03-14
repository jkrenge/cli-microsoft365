var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListWebhookSetCommand_instances, _SpoListWebhookSetCommand_initTelemetry, _SpoListWebhookSetCommand_initOptions, _SpoListWebhookSetCommand_initValidators, _SpoListWebhookSetCommand_initOptionSets;
import chalk from 'chalk';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListWebhookSetCommand extends SpoCommand {
    get name() {
        return commands.LIST_WEBHOOK_SET;
    }
    get description() {
        return 'Updates the specified webhook';
    }
    constructor() {
        super();
        _SpoListWebhookSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListWebhookSetCommand_instances, "m", _SpoListWebhookSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookSetCommand_instances, "m", _SpoListWebhookSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookSetCommand_instances, "m", _SpoListWebhookSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListWebhookSetCommand_instances, "m", _SpoListWebhookSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Updating webhook ${args.options.id} belonging to list ${args.options.listId || args.options.listTitle || args.options.listUrl} located at site ${args.options.webUrl}...`);
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
        const requestBody = {
            notificationUrl: args.options.notificationUrl,
            expirationDateTime: args.options.expirationDateTime,
            clientState: args.options.clientState
        };
        const requestOptions = {
            url: requestUrl,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            data: requestBody,
            responseType: 'json'
        };
        try {
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListWebhookSetCommand_instances = new WeakSet(), _SpoListWebhookSetCommand_initTelemetry = function _SpoListWebhookSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            notificationUrl: typeof args.options.notificationUrl !== 'undefined',
            expirationDateTime: typeof args.options.expirationDateTime !== 'undefined',
            clientState: typeof args.options.clientState !== 'undefined'
        });
    });
}, _SpoListWebhookSetCommand_initOptions = function _SpoListWebhookSetCommand_initOptions() {
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
        option: '-n, --notificationUrl [notificationUrl]'
    }, {
        option: '-e, --expirationDateTime [expirationDateTime]'
    }, {
        option: '-c, --clientState [clientState]'
    });
}, _SpoListWebhookSetCommand_initValidators = function _SpoListWebhookSetCommand_initValidators() {
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
        if (!args.options.notificationUrl && !args.options.expirationDateTime && !args.options.clientState) {
            return 'Specify notificationUrl, expirationDateTime, clientState or multiple, at least one is required';
        }
        const parsedDateTime = Date.parse(args.options.expirationDateTime);
        if (args.options.expirationDateTime && !(!parsedDateTime) !== true) {
            return `${args.options.expirationDateTime} is not a valid date format. Provide the date in one of the following formats:
      ${chalk.grey('YYYY-MM-DD')}
      ${chalk.grey('YYYY-MM-DDThh:mm')}
      ${chalk.grey('YYYY-MM-DDThh:mmZ')}
      ${chalk.grey('YYYY-MM-DDThh:mm±hh:mm')}`;
        }
        return true;
    });
}, _SpoListWebhookSetCommand_initOptionSets = function _SpoListWebhookSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListWebhookSetCommand();
//# sourceMappingURL=list-webhook-set.js.map