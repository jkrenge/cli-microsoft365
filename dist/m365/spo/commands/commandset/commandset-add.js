var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCommandSetAddCommand_instances, _a, _SpoCommandSetAddCommand_initTelemetry, _SpoCommandSetAddCommand_initOptions, _SpoCommandSetAddCommand_initValidators;
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import request from '../../../../request.js';
class SpoCommandSetAddCommand extends SpoCommand {
    get name() {
        return commands.COMMANDSET_ADD;
    }
    get description() {
        return 'Add a ListView Command Set to a site.';
    }
    constructor() {
        super();
        _SpoCommandSetAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCommandSetAddCommand_instances, "m", _SpoCommandSetAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetAddCommand_instances, "m", _SpoCommandSetAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetAddCommand_instances, "m", _SpoCommandSetAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding ListView Command Set ${args.options.clientSideComponentId} to site '${args.options.webUrl}'...`);
        }
        if (!args.options.scope) {
            args.options.scope = 'Web';
        }
        const location = args.options.location && this.getLocation(args.options.location);
        const listType = this.getListTemplate(args.options.listType);
        try {
            const requestBody = {
                Title: args.options.title,
                Description: args.options.description,
                Location: location,
                ClientSideComponentId: args.options.clientSideComponentId,
                RegistrationId: listType,
                RegistrationType: 1
            };
            if (args.options.clientSideComponentProperties) {
                requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
            }
            const requestOptions = {
                url: `${args.options.webUrl}/_api/${args.options.scope}/UserCustomActions`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                data: requestBody,
                responseType: 'json'
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
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
_a = SpoCommandSetAddCommand, _SpoCommandSetAddCommand_instances = new WeakSet(), _SpoCommandSetAddCommand_initTelemetry = function _SpoCommandSetAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            location: typeof args.options.location !== 'undefined'
        });
    });
}, _SpoCommandSetAddCommand_initOptions = function _SpoCommandSetAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listType <listType>', autocomplete: _a.listTypes
    }, {
        option: '-i, --clientSideComponentId  <clientSideComponentId>'
    }, {
        option: '--description [description]'
    }, {
        option: '--clientSideComponentProperties  [clientSideComponentProperties]'
    }, {
        option: '-s, --scope [scope]', autocomplete: _a.scopes
    }, {
        option: '--location [location]', autocomplete: _a.locations
    });
}, _SpoCommandSetAddCommand_initValidators = function _SpoCommandSetAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        if (_a.listTypes.indexOf(args.options.listType) < 0) {
            return `${args.options.listType} is not a valid list type. Allowed values are ${_a.listTypes.join(', ')}`;
        }
        if (args.options.scope && _a.scopes.indexOf(args.options.scope) < 0) {
            return `${args.options.scope} is not a valid scope. Allowed values are ${_a.scopes.join(', ')}`;
        }
        if (args.options.location && _a.locations.indexOf(args.options.location) < 0) {
            return `${args.options.location} is not a valid location. Allowed values are ${_a.locations.join(', ')}`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
SpoCommandSetAddCommand.listTypes = ['List', 'Library', 'SitePages'];
SpoCommandSetAddCommand.scopes = ['Site', 'Web'];
SpoCommandSetAddCommand.locations = ['ContextMenu', 'CommandBar', 'Both'];
export default new SpoCommandSetAddCommand();
//# sourceMappingURL=commandset-add.js.map