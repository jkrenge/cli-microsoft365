var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCustomActionListCommand_instances, _SpoCustomActionListCommand_initTelemetry, _SpoCustomActionListCommand_initOptions, _SpoCustomActionListCommand_initValidators;
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCustomActionListCommand extends SpoCommand {
    get name() {
        return commands.CUSTOMACTION_LIST;
    }
    get description() {
        return 'Lists all user custom actions at the given scope';
    }
    defaultProperties() {
        return ['Name', 'Location', 'Scope', 'Id'];
    }
    constructor() {
        super();
        _SpoCustomActionListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCustomActionListCommand_instances, "m", _SpoCustomActionListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionListCommand_instances, "m", _SpoCustomActionListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCustomActionListCommand_instances, "m", _SpoCustomActionListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const scope = args.options.scope ? args.options.scope : 'All';
            if (this.debug) {
                await logger.logToStderr(`Attempt to get custom actions list with scope: ${scope}`);
                await logger.logToStderr('');
            }
            const customActions = await spo.getCustomActions(args.options.webUrl, args.options.scope);
            if (args.options.output !== 'json') {
                customActions.forEach(a => a.Scope = this.humanizeScope(a.Scope));
            }
            await logger.log(customActions);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    humanizeScope(scope) {
        switch (scope) {
            case 2:
                return "Site";
            case 3:
                return "Web";
        }
        return `${scope}`;
    }
}
_SpoCustomActionListCommand_instances = new WeakSet(), _SpoCustomActionListCommand_initTelemetry = function _SpoCustomActionListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            scope: args.options.scope || 'All'
        });
    });
}, _SpoCustomActionListCommand_initOptions = function _SpoCustomActionListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web', 'All']
    });
}, _SpoCustomActionListCommand_initValidators = function _SpoCustomActionListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (typeof isValidUrl === 'string') {
            return isValidUrl;
        }
        if (args.options.scope) {
            if (args.options.scope !== 'Site' &&
                args.options.scope !== 'Web' &&
                args.options.scope !== 'All') {
                return `${args.options.scope} is not a valid custom action scope. Allowed values are Site|Web|All`;
            }
        }
        return true;
    });
};
export default new SpoCustomActionListCommand();
//# sourceMappingURL=customaction-list.js.map