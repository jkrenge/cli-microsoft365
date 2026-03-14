var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoEventreceiverGetCommand_instances, _SpoEventreceiverGetCommand_initTelemetry, _SpoEventreceiverGetCommand_initOptions, _SpoEventreceiverGetCommand_initValidators, _SpoEventreceiverGetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoEventreceiverGetCommand extends SpoCommand {
    get name() {
        return commands.EVENTRECEIVER_GET;
    }
    get description() {
        return 'Gets a specific event receiver attached to the web, site or list (if any of the list options are filled in) by receiver name of id';
    }
    constructor() {
        super();
        _SpoEventreceiverGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoEventreceiverGetCommand_instances, "m", _SpoEventreceiverGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoEventreceiverGetCommand_instances, "m", _SpoEventreceiverGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoEventreceiverGetCommand_instances, "m", _SpoEventreceiverGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoEventreceiverGetCommand_instances, "m", _SpoEventreceiverGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const eventReceiver = await this.getEventReceiver(args);
            await logger.log(eventReceiver);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getEventReceiver(args) {
        const requestOptions = {
            url: `${args.options.webUrl}/_api`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        if (!args.options.scope || args.options.scope === 'web') {
            requestOptions.url += `/web`;
        }
        else {
            requestOptions.url += '/site';
        }
        if (args.options.listId) {
            requestOptions.url += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            requestOptions.url += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            requestOptions.url += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        requestOptions.url += '/eventreceivers';
        if (args.options.id) {
            requestOptions.url += `(guid'${args.options.id}')`;
            const res = await request.get(requestOptions);
            // Instead of failing, the request will return an 'odata.null' object if the event receiver doesn't exist.
            if (res && res['odata.null']) {
                throw `The specified event receiver '${args.options.id}' does not exist.`;
            }
            return res;
        }
        else {
            requestOptions.url += `?$filter=receivername eq '${args.options.name}'`;
            const res = await request.get(requestOptions);
            if (res.value.length === 0) {
                throw `The specified event receiver '${args.options.name}' does not exist.`;
            }
            if (res.value && res.value.length > 1) {
                const resultAsKeyValuePair = formatting.convertArrayToHashTable('ReceiverId', res.value);
                return await cli.handleMultipleResultsFound(`Multiple event receivers with name '${args.options.name}' found.`, resultAsKeyValuePair);
            }
            return res.value[0];
        }
    }
}
_SpoEventreceiverGetCommand_instances = new WeakSet(), _SpoEventreceiverGetCommand_initTelemetry = function _SpoEventreceiverGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _SpoEventreceiverGetCommand_initOptions = function _SpoEventreceiverGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listId  [listId]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: ['web', 'site']
    });
}, _SpoEventreceiverGetCommand_initValidators = function _SpoEventreceiverGetCommand_initValidators() {
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
}, _SpoEventreceiverGetCommand_initOptionSets = function _SpoEventreceiverGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['name', 'id'] });
};
export default new SpoEventreceiverGetCommand();
//# sourceMappingURL=eventreceiver-get.js.map