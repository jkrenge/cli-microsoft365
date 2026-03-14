var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCommandSetSetCommand_instances, _a, _SpoCommandSetSetCommand_initTelemetry, _SpoCommandSetSetCommand_initOptions, _SpoCommandSetSetCommand_initValidators, _SpoCommandSetSetCommand_initOptionSets;
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { cli } from '../../../../cli/cli.js';
class SpoCommandSetSetCommand extends SpoCommand {
    get name() {
        return commands.COMMANDSET_SET;
    }
    get description() {
        return 'Updates a ListView Command Set on a site.';
    }
    constructor() {
        super();
        _SpoCommandSetSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCommandSetSetCommand_instances, "m", _SpoCommandSetSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetSetCommand_instances, "m", _SpoCommandSetSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetSetCommand_instances, "m", _SpoCommandSetSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetSetCommand_instances, "m", _SpoCommandSetSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Updating ListView Command Set ${args.options.id || args.options.title || args.options.clientSideComponentId} to site '${args.options.webUrl}'...`);
        }
        if (!args.options.scope) {
            args.options.scope = 'All';
        }
        const location = this.getLocation(args.options.location ? args.options.location : '');
        try {
            const requestBody = {};
            if (args.options.newTitle) {
                requestBody.Title = args.options.newTitle;
            }
            if (args.options.description !== undefined) {
                requestBody.Description = args.options.description;
            }
            if (args.options.location) {
                requestBody.Location = location;
            }
            if (args.options.listType) {
                requestBody.RegistrationId = this.getListTemplate(args.options.listType);
            }
            if (args.options.clientSideComponentProperties) {
                requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
            }
            if (args.options.newClientSideComponentId) {
                requestBody.ClientSideComponentId = args.options.newClientSideComponentId;
            }
            const customAction = await this.getCustomAction(args.options);
            const requestOptions = {
                url: `${args.options.webUrl}/_api/${customAction.Scope === 3 ? "Web" : "Site"}/UserCustomActions('${formatting.encodeQueryParameter(customAction.Id)}')`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'X-HTTP-Method': 'MERGE'
                },
                data: requestBody,
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
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
    getLocation(location) {
        switch (location) {
            case 'Both':
                return 'ClientSideExtension.ListViewCommandSet';
            case 'ContextMenu':
                return 'ClientSideExtension.ListViewCommandSet.ContextMenu';
            default:
                return 'ClientSideExtension.ListViewCommandSet.CommandBar';
        }
    }
    getListTemplate(listTemplate) {
        switch (listTemplate) {
            case 'SitePages':
                return '119';
            case 'Library':
                return '101';
            default:
                return '100';
        }
    }
}
_a = SpoCommandSetSetCommand, _SpoCommandSetSetCommand_instances = new WeakSet(), _SpoCommandSetSetCommand_initTelemetry = function _SpoCommandSetSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            newClientSideComponentId: typeof args.options.newClientSideComponentId !== 'undefined',
            newTitle: typeof args.options.newTitle !== 'undefined',
            description: typeof args.options.description !== 'undefined',
            listType: typeof args.options.listType !== 'undefined',
            clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            location: typeof args.options.location !== 'undefined'
        });
    });
}, _SpoCommandSetSetCommand_initOptions = function _SpoCommandSetSetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
    }, {
        option: '--newClientSideComponentId  [newClientSideComponentId]'
    }, {
        option: '--newTitle [newTitle]'
    }, {
        option: '--description [description]'
    }, {
        option: '-l, --listType [listType]', autocomplete: _a.listTypes
    }, {
        option: '--clientSideComponentProperties  [clientSideComponentProperties]'
    }, {
        option: '-s, --scope [scope]', autocomplete: _a.scopes
    }, {
        option: '--location [location]', autocomplete: _a.locations
    });
}, _SpoCommandSetSetCommand_initValidators = function _SpoCommandSetSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        if (args.options.newClientSideComponentId && !validation.isValidGuid(args.options.newClientSideComponentId)) {
            return `${args.options.newClientSideComponentId} is not a valid GUID`;
        }
        if (args.options.listType && _a.listTypes.indexOf(args.options.listType) < 0) {
            return `${args.options.listType} is not a valid list type. Allowed values are ${_a.listTypes.join(', ')}`;
        }
        if (args.options.scope && _a.scopes.indexOf(args.options.scope) < 0) {
            return `${args.options.scope} is not a valid scope. Allowed values are ${_a.scopes.join(', ')}`;
        }
        if (args.options.location && _a.locations.indexOf(args.options.location) < 0) {
            return `${args.options.location} is not a valid location. Allowed values are ${_a.locations.join(', ')}`;
        }
        if (!args.options.newTitle && args.options.description === undefined && !args.options.listType && !args.options.clientSideComponentProperties && !args.options.location && !args.options.newClientSideComponentId) {
            return `Please specify option to be updated`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoCommandSetSetCommand_initOptionSets = function _SpoCommandSetSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'clientSideComponentId'] });
};
SpoCommandSetSetCommand.listTypes = ['List', 'Library', 'SitePages'];
SpoCommandSetSetCommand.scopes = ['All', 'Site', 'Web'];
SpoCommandSetSetCommand.locations = ['ContextMenu', 'CommandBar', 'Both'];
export default new SpoCommandSetSetCommand();
//# sourceMappingURL=commandset-set.js.map