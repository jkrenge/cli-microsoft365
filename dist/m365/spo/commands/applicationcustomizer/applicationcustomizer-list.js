var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoApplicationCustomizerListCommand_instances, _a, _SpoApplicationCustomizerListCommand_initOptions, _SpoApplicationCustomizerListCommand_initTelemetry, _SpoApplicationCustomizerListCommand_initValidators;
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoApplicationCustomizerListCommand extends SpoCommand {
    get name() {
        return commands.APPLICATIONCUSTOMIZER_LIST;
    }
    get description() {
        return 'Get a list of application customizers that are added to a site.';
    }
    defaultProperties() {
        return ['Name', 'Location', 'Scope', 'Id'];
    }
    constructor() {
        super();
        _SpoApplicationCustomizerListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerListCommand_instances, "m", _SpoApplicationCustomizerListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerListCommand_instances, "m", _SpoApplicationCustomizerListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoApplicationCustomizerListCommand_instances, "m", _SpoApplicationCustomizerListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving application customizers...`);
        }
        const applicationCustomizers = await spo.getCustomActions(args.options.webUrl, args.options.scope, `Location eq 'ClientSideExtension.ApplicationCustomizer'`);
        await logger.log(applicationCustomizers);
    }
}
_a = SpoApplicationCustomizerListCommand, _SpoApplicationCustomizerListCommand_instances = new WeakSet(), _SpoApplicationCustomizerListCommand_initOptions = function _SpoApplicationCustomizerListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: _a.scopes
    });
}, _SpoApplicationCustomizerListCommand_initTelemetry = function _SpoApplicationCustomizerListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            scope: typeof args.options.scope !== 'undefined'
        });
    });
}, _SpoApplicationCustomizerListCommand_initValidators = function _SpoApplicationCustomizerListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.scope && _a.scopes.indexOf(args.options.scope) < 0) {
            return `${args.options.scope} is not a valid scope. Allowed values are ${_a.scopes.join(', ')}`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
SpoApplicationCustomizerListCommand.scopes = ['All', 'Site', 'Web'];
export default new SpoApplicationCustomizerListCommand();
//# sourceMappingURL=applicationcustomizer-list.js.map