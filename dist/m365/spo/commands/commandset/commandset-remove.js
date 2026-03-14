var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCommandSetRemoveCommand_instances, _a, _SpoCommandSetRemoveCommand_initTelemetry, _SpoCommandSetRemoveCommand_initOptions, _SpoCommandSetRemoveCommand_initValidators, _SpoCommandSetRemoveCommand_initOptionSets;
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { cli } from '../../../../cli/cli.js';
class SpoCommandSetRemoveCommand extends SpoCommand {
    get name() {
        return commands.COMMANDSET_REMOVE;
    }
    get description() {
        return 'Remove a ListView Command Set that is added to a site.';
    }
    constructor() {
        super();
        _SpoCommandSetRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCommandSetRemoveCommand_instances, "m", _SpoCommandSetRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetRemoveCommand_instances, "m", _SpoCommandSetRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetRemoveCommand_instances, "m", _SpoCommandSetRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetRemoveCommand_instances, "m", _SpoCommandSetRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing ListView Command Set ${args.options.clientSideComponentId || args.options.title || args.options.id} to site '${args.options.webUrl}'...`);
        }
        if (args.options.force) {
            await this.deleteCommandset(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove command set '${args.options.clientSideComponentId || args.options.title || args.options.id}'?` });
            if (result) {
                await this.deleteCommandset(args);
            }
        }
    }
    async getCustomAction(options) {
        let commandSets = [];
        if (options.id) {
            const commandSet = await spo.getCustomActionById(options.webUrl, options.id, options.scope);
            if (commandSet) {
                commandSets.push(commandSet);
            }
        }
        else if (options.title) {
            commandSets = await spo.getCustomActions(options.webUrl, options.scope, `(Title eq '${formatting.encodeQueryParameter(options.title)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`);
        }
        else {
            commandSets = await spo.getCustomActions(options.webUrl, options.scope, `(ClientSideComponentId eq guid'${options.clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`);
        }
        if (commandSets.length === 0) {
            throw `No user commandsets with ${options.title && `title '${options.title}'` || options.clientSideComponentId && `ClientSideComponentId '${options.clientSideComponentId}'` || options.id && `id '${options.id}'`} found`;
        }
        if (commandSets.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', commandSets);
            return await cli.handleMultipleResultsFound(`Multiple user commandsets with ${options.title ? `title '${options.title}'` : `ClientSideComponentId '${options.clientSideComponentId}'`} found.`, resultAsKeyValuePair);
        }
        return commandSets[0];
    }
    async deleteCommandset(args) {
        if (!args.options.scope) {
            args.options.scope = 'All';
        }
        try {
            const customAction = await this.getCustomAction(args.options);
            const requestOptions = {
                url: `${args.options.webUrl}/_api/${customAction.Scope === 3 ? "Web" : "Site"}/UserCustomActions('${formatting.encodeQueryParameter(customAction.Id)}')`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = SpoCommandSetRemoveCommand, _SpoCommandSetRemoveCommand_instances = new WeakSet(), _SpoCommandSetRemoveCommand_initTelemetry = function _SpoCommandSetRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoCommandSetRemoveCommand_initOptions = function _SpoCommandSetRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
    }, {
        option: '-s, --scope [scope]', autocomplete: _a.scopes
    }, {
        option: '-f, --force'
    });
}, _SpoCommandSetRemoveCommand_initValidators = function _SpoCommandSetRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        if (args.options.scope && _a.scopes.indexOf(args.options.scope) < 0) {
            return `${args.options.scope} is not a valid scope. Allowed values are ${_a.scopes.join(', ')}`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoCommandSetRemoveCommand_initOptionSets = function _SpoCommandSetRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'clientSideComponentId'] });
};
SpoCommandSetRemoveCommand.scopes = ['All', 'Site', 'Web'];
export default new SpoCommandSetRemoveCommand();
//# sourceMappingURL=commandset-remove.js.map