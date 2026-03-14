var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoEventreceiverRemoveCommand_instances, _SpoEventreceiverRemoveCommand_initTelemetry, _SpoEventreceiverRemoveCommand_initOptions, _SpoEventreceiverRemoveCommand_initValidators, _SpoEventreceiverRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import getCommand from './eventreceiver-get.js';
class SpoEventreceiverRemoveCommand extends SpoCommand {
    get name() {
        return commands.EVENTRECEIVER_REMOVE;
    }
    get description() {
        return 'Removes event receivers for the specified web, site, or list.';
    }
    constructor() {
        super();
        _SpoEventreceiverRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoEventreceiverRemoveCommand_instances, "m", _SpoEventreceiverRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoEventreceiverRemoveCommand_instances, "m", _SpoEventreceiverRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoEventreceiverRemoveCommand_instances, "m", _SpoEventreceiverRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoEventreceiverRemoveCommand_instances, "m", _SpoEventreceiverRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeEventReceiver(args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove event receiver with ${args.options.id ? `id ${args.options.id}` : `name ${args.options.name}`}?` });
            if (result) {
                await this.removeEventReceiver(args.options);
            }
        }
    }
    async removeEventReceiver(options) {
        try {
            let requestUrl = `${options.webUrl}/_api/${options.scope || 'web'}`;
            if (options.listId) {
                requestUrl += `/lists('${options.listId}')`;
            }
            else if (options.listTitle) {
                requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
            }
            else if (options.listUrl) {
                const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
                requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
            }
            const rerId = await this.getEventReceiverId(options);
            requestUrl += `/eventreceivers('${rerId}')`;
            const requestOptions = {
                url: requestUrl,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getEventReceiverId(options) {
        if (options.id) {
            return options.id;
        }
        const getOptions = {
            webUrl: options.webUrl,
            listId: options.listId,
            listTitle: options.listTitle,
            listUrl: options.listUrl,
            scope: options.scope,
            id: options.id,
            name: options.name,
            debug: this.debug,
            verbose: this.verbose
        };
        const commandOutput = await cli.executeCommandWithOutput(getCommand, { options: { ...getOptions, _: [] } });
        const eventReceiver = JSON.parse(commandOutput.stdout);
        return eventReceiver.ReceiverId;
    }
}
_SpoEventreceiverRemoveCommand_instances = new WeakSet(), _SpoEventreceiverRemoveCommand_initTelemetry = function _SpoEventreceiverRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            scope: args.options.scope || 'web',
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoEventreceiverRemoveCommand_initOptions = function _SpoEventreceiverRemoveCommand_initOptions() {
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
    }, {
        option: '-f, --force'
    });
}, _SpoEventreceiverRemoveCommand_initValidators = function _SpoEventreceiverRemoveCommand_initValidators() {
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
}, _SpoEventreceiverRemoveCommand_initOptionSets = function _SpoEventreceiverRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['name', 'id'] });
};
export default new SpoEventreceiverRemoveCommand();
//# sourceMappingURL=eventreceiver-remove.js.map