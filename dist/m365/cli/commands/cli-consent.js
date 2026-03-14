var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _CliConsentCommand_instances, _CliConsentCommand_initTelemetry, _CliConsentCommand_initOptions, _CliConsentCommand_initValidators;
import { cli } from '../../../cli/cli.js';
import AnonymousCommand from '../../base/AnonymousCommand.js';
import commands from '../commands.js';
class CliConsentCommand extends AnonymousCommand {
    get name() {
        return commands.CONSENT;
    }
    get description() {
        return 'Consent additional permissions for the Microsoft Entra application used by the CLI for Microsoft 365';
    }
    constructor() {
        super();
        _CliConsentCommand_instances.add(this);
        __classPrivateFieldGet(this, _CliConsentCommand_instances, "m", _CliConsentCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _CliConsentCommand_instances, "m", _CliConsentCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _CliConsentCommand_instances, "m", _CliConsentCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let scope = '';
        if (args.options.service === 'VivaEngage') {
            scope = 'https://api.yammer.com/user_impersonation';
        }
        await logger.log(`To consent permissions for executing ${args.options.service} commands, navigate in your web browser to https://login.microsoftonline.com/${cli.getTenant()}/oauth2/v2.0/authorize?client_id=${cli.getClientId()}&response_type=code&scope=${encodeURIComponent(scope)}`);
    }
    async action(logger, args) {
        await this.initAction(args, logger);
        await this.commandAction(logger, args);
    }
}
_CliConsentCommand_instances = new WeakSet(), _CliConsentCommand_initTelemetry = function _CliConsentCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            service: args.options.service
        });
    });
}, _CliConsentCommand_initOptions = function _CliConsentCommand_initOptions() {
    this.options.unshift({
        option: '-s, --service <service>',
        autocomplete: ['VivaEngage']
    });
}, _CliConsentCommand_initValidators = function _CliConsentCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.service !== 'VivaEngage') {
            return `${args.options.service} is not a valid value for the service option. Allowed values: VivaEngage`;
        }
        return true;
    });
};
export default new CliConsentCommand();
//# sourceMappingURL=cli-consent.js.map