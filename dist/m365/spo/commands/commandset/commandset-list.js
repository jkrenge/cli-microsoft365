var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCommandSetListCommand_instances, _a, _SpoCommandSetListCommand_initTelemetry, _SpoCommandSetListCommand_initOptions, _SpoCommandSetListCommand_initValidators;
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCommandSetListCommand extends SpoCommand {
    get name() {
        return commands.COMMANDSET_LIST;
    }
    get description() {
        return 'Get a list of ListView Command Sets that are added to a site.';
    }
    defaultProperties() {
        return ['Name', 'Location', 'Scope', 'Id'];
    }
    constructor() {
        super();
        _SpoCommandSetListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCommandSetListCommand_instances, "m", _SpoCommandSetListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetListCommand_instances, "m", _SpoCommandSetListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCommandSetListCommand_instances, "m", _SpoCommandSetListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Attempt to get commandsets...`);
            }
            const commandsets = await spo.getCustomActions(args.options.webUrl, args.options.scope, `startswith(Location,'ClientSideExtension.ListViewCommandSet')`);
            await logger.log(commandsets);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_a = SpoCommandSetListCommand, _SpoCommandSetListCommand_instances = new WeakSet(), _SpoCommandSetListCommand_initTelemetry = function _SpoCommandSetListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            scope: args.options.scope || 'All'
        });
    });
}, _SpoCommandSetListCommand_initOptions = function _SpoCommandSetListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: _a.scopes
    });
}, _SpoCommandSetListCommand_initValidators = function _SpoCommandSetListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.scope && _a.scopes.indexOf(args.options.scope) < 0) {
            return `${args.options.scope} is not a valid scope. Valid scopes are ${_a.scopes.join(', ')}`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
SpoCommandSetListCommand.scopes = ['All', 'Site', 'Web'];
export default new SpoCommandSetListCommand();
//# sourceMappingURL=commandset-list.js.map