var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoEventreceiverListCommand_instances, _SpoEventreceiverListCommand_initTelemetry, _SpoEventreceiverListCommand_initOptions, _SpoEventreceiverListCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoEventreceiverListCommand extends SpoCommand {
    get name() {
        return commands.EVENTRECEIVER_LIST;
    }
    get description() {
        return 'Retrieves event receivers for the specified web, site or list.';
    }
    defaultProperties() {
        return ['ReceiverId', 'ReceiverName'];
    }
    constructor() {
        super();
        _SpoEventreceiverListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoEventreceiverListCommand_instances, "m", _SpoEventreceiverListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoEventreceiverListCommand_instances, "m", _SpoEventreceiverListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoEventreceiverListCommand_instances, "m", _SpoEventreceiverListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let requestUrl = `${args.options.webUrl}/_api/`;
        let listUrl = '';
        if (args.options.listId) {
            listUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
        }
        else if (args.options.listTitle) {
            listUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            listUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
        }
        if (!args.options.scope || args.options.scope === 'web') {
            requestUrl += `web/${listUrl}eventreceivers`;
        }
        else {
            requestUrl += 'site/eventreceivers';
        }
        try {
            const res = await odata.getAllItems(requestUrl);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoEventreceiverListCommand_instances = new WeakSet(), _SpoEventreceiverListCommand_initTelemetry = function _SpoEventreceiverListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            scope: typeof args.options.scope !== 'undefined'
        });
    });
}, _SpoEventreceiverListCommand_initOptions = function _SpoEventreceiverListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listId  [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: ['web', 'site']
    });
}, _SpoEventreceiverListCommand_initValidators = function _SpoEventreceiverListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        const listOptions = [args.options.listId, args.options.listTitle, args.options.listUrl];
        if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
            return `Specify either list id or title or list url`;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        if (args.options.scope && ['web', 'site'].indexOf(args.options.scope) === -1) {
            return `${args.options.scope} is not a valid type value. Allowed values web|site.`;
        }
        if (args.options.scope && args.options.scope === 'site' && (args.options.listId || args.options.listUrl || args.options.listTitle)) {
            return 'Scope cannot be set to site when retrieving list event receivers.';
        }
        return true;
    });
};
export default new SpoEventreceiverListCommand();
//# sourceMappingURL=eventreceiver-list.js.map