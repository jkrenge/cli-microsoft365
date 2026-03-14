var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoApplicationCustomizerAddCommand_instances, _a, _SpoApplicationCustomizerAddCommand_initOptions, _SpoApplicationCustomizerAddCommand_initTelemetry, _SpoApplicationCustomizerAddCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoApplicationCustomizerAddCommand extends SpoCommand {
    get name() {
        return commands.APPLICATIONCUSTOMIZER_ADD;
    }
    get description() {
        return 'Add an application customizer to a site.';
    }
    constructor() {
        super();
        _SpoApplicationCustomizerAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerAddCommand_instances, "m", _SpoApplicationCustomizerAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerAddCommand_instances, "m", _SpoApplicationCustomizerAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerAddCommand_instances, "m", _SpoApplicationCustomizerAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding application customizer with title '${args.options.title}' and clientSideComponentId '${args.options.clientSideComponentId}' to the site`);
        }
        const requestBody = {
            Title: args.options.title,
            Name: args.options.title,
            Description: args.options.description,
            Location: 'ClientSideExtension.ApplicationCustomizer',
            ClientSideComponentId: args.options.clientSideComponentId
        };
        if (args.options.clientSideComponentProperties) {
            requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
        }
        const scope = args.options.scope || 'Site';
        const requestOptions = {
            url: `${args.options.webUrl}/_api/${scope}/UserCustomActions`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: requestBody,
            responseType: 'json'
        };
        await request.post(requestOptions);
    }
}
_a = SpoApplicationCustomizerAddCommand, _SpoApplicationCustomizerAddCommand_instances = new WeakSet(), _SpoApplicationCustomizerAddCommand_initOptions = function _SpoApplicationCustomizerAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
    }, {
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --clientSideComponentId <clientSideComponentId>'
    }, {
        option: '--description [description]'
    }, {
        option: '--clientSideComponentProperties [clientSideComponentProperties]'
    }, {
        option: '-s, --scope [scope]', autocomplete: _a.scopes
    });
}, _SpoApplicationCustomizerAddCommand_initTelemetry = function _SpoApplicationCustomizerAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
            scope: typeof args.options.scope !== 'undefined'
        });
    });
}, _SpoApplicationCustomizerAddCommand_initValidators = function _SpoApplicationCustomizerAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.webUrl) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        if (!validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        if (args.options.clientSideComponentProperties) {
            try {
                JSON.parse(args.options.clientSideComponentProperties);
            }
            catch (e) {
                return `An error has occurred while parsing clientSideComponentProperties: ${e}`;
            }
        }
        if (args.options.scope && _a.scopes.indexOf(args.options.scope) < 0) {
            return `${args.options.scope} is not a valid value for allowedMembers. Valid values are ${_a.scopes.join(', ')}`;
        }
        return true;
    });
};
SpoApplicationCustomizerAddCommand.scopes = ['Site', 'Web'];
export default new SpoApplicationCustomizerAddCommand();
//# sourceMappingURL=applicationcustomizer-add.js.map