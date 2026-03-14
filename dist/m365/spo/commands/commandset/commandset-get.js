var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCommandSetGetCommand_instances, _a, _SpoCommandSetGetCommand_initTelemetry, _SpoCommandSetGetCommand_initOptions, _SpoCommandSetGetCommand_initValidators, _SpoCommandSetGetCommand_initOptionSets, _SpoCommandSetGetCommand_initTypes;
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class SpoCommandSetGetCommand extends SpoCommand {
    get name() {
        return commands.COMMANDSET_GET;
    }
    get description() {
        return 'Get a ListView Command Set that is added to a site.';
    }
    constructor() {
        super();
        _SpoCommandSetGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCommandSetGetCommand_instances, "m", _SpoCommandSetGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetGetCommand_instances, "m", _SpoCommandSetGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetGetCommand_instances, "m", _SpoCommandSetGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetGetCommand_instances, "m", _SpoCommandSetGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetGetCommand_instances, "m", _SpoCommandSetGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Attempt to get a specific Command Set by property ${args.options.title || args.options.id || args.options.clientSideComponentId}.`);
            }
            let commandSet;
            if (args.options.id) {
                const customAction = await spo.getCustomActionById(args.options.webUrl, args.options.id, args.options.scope);
                if (customAction === undefined) {
                    throw `Command set with id ${args.options.id} can't be found.`;
                }
                else if (!_a.allowedCommandSetLocations.some(allowedLocation => allowedLocation === customAction.Location)) {
                    throw `Custom action with id ${args.options.id} is not a command set.`;
                }
                commandSet = customAction;
            }
            else if (args.options.clientSideComponentId) {
                const filter = `${this.getBaseFilter()} ClientSideComponentId eq guid'${args.options.clientSideComponentId}'`;
                const customActions = await spo.getCustomActions(args.options.webUrl, args.options.scope, filter);
                if (customActions.length === 0) {
                    throw `No command set with clientSideComponentId '${args.options.clientSideComponentId}' found.`;
                }
                commandSet = customActions[0];
            }
            else if (args.options.title) {
                const filter = `${this.getBaseFilter()} Title eq '${formatting.encodeQueryParameter(args.options.title)}'`;
                const customActions = await spo.getCustomActions(args.options.webUrl, args.options.scope, filter);
                if (customActions.length === 1) {
                    commandSet = customActions[0];
                }
                else if (customActions.length === 0) {
                    throw `No command set with title '${args.options.title}' found.`;
                }
                else {
                    const resultAsKeyValuePair = formatting.convertArrayToHashTable('Id', customActions);
                    commandSet = await cli.handleMultipleResultsFound(`Multiple command sets with title '${args.options.title}' found.`, resultAsKeyValuePair);
                }
            }
            if (!args.options.clientSideComponentProperties) {
                await logger.log(commandSet);
            }
            else {
                const properties = formatting.tryParseJson(commandSet.ClientSideComponentProperties);
                await logger.log(properties);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getBaseFilter() {
        return `startswith(Location,'${_a.baseLocation}') and`;
    }
}
_a = SpoCommandSetGetCommand, _SpoCommandSetGetCommand_instances = new WeakSet(), _SpoCommandSetGetCommand_initTelemetry = function _SpoCommandSetGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            clientSideComponentProperties: !!args.options.clientSideComponentProperties
        });
    });
}, _SpoCommandSetGetCommand_initOptions = function _SpoCommandSetGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-c, --clientSideComponentId [clientSideComponentId]'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: _a.scopes
    }, {
        option: '-p, --clientSideComponentProperties'
    });
}, _SpoCommandSetGetCommand_initValidators = function _SpoCommandSetGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID.`;
        }
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID.`;
        }
        if (args.options.scope && _a.scopes.indexOf(args.options.scope) < 0) {
            return `${args.options.scope} is not a valid scope. Valid scopes are ${_a.scopes.join(', ')}`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoCommandSetGetCommand_initOptionSets = function _SpoCommandSetGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['title', 'id', 'clientSideComponentId'] });
}, _SpoCommandSetGetCommand_initTypes = function _SpoCommandSetGetCommand_initTypes() {
    this.types.string.push('webUrl', 'title', 'id', 'clientSideComponentId', 'scope');
    this.types.boolean.push('clientSideComponentProperties');
};
SpoCommandSetGetCommand.scopes = ['All', 'Site', 'Web'];
SpoCommandSetGetCommand.baseLocation = 'ClientSideExtension.ListViewCommandSet';
SpoCommandSetGetCommand.allowedCommandSetLocations = [_a.baseLocation, `${_a.baseLocation}.CommandBar`, `${_a.baseLocation}.ContextMenu`];
export default new SpoCommandSetGetCommand();
//# sourceMappingURL=commandset-get.js.map