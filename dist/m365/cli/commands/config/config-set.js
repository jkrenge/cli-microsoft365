var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _CliConfigSetCommand_instances, _a, _CliConfigSetCommand_initTelemetry, _CliConfigSetCommand_initOptions, _CliConfigSetCommand_initValidators;
import { AuthType } from "../../../../Auth.js";
import { cli } from "../../../../cli/cli.js";
import { settingsNames } from "../../../../settingsNames.js";
import { validation } from "../../../../utils/validation.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";
class CliConfigSetCommand extends AnonymousCommand {
    get name() {
        return commands.CONFIG_SET;
    }
    get description() {
        return 'Manage global configuration settings about the CLI for Microsoft 365';
    }
    constructor() {
        super();
        _CliConfigSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _CliConfigSetCommand_instances, "m", _CliConfigSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _CliConfigSetCommand_instances, "m", _CliConfigSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _CliConfigSetCommand_instances, "m", _CliConfigSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let value;
        switch (args.options.key) {
            case settingsNames.autoOpenLinksInBrowser:
            case settingsNames.copyDeviceCodeToClipboard:
            case settingsNames.csvHeader:
            case settingsNames.csvQuoted:
            case settingsNames.csvQuotedEmpty:
            case settingsNames.disableTelemetry:
            case settingsNames.printErrorsAsPlainText:
            case settingsNames.prompt:
            case settingsNames.showHelpOnFailure:
                value = args.options.value === 'true';
                break;
            default:
                value = args.options.value;
                break;
        }
        cli.getConfig().set(args.options.key, value);
    }
}
_a = CliConfigSetCommand, _CliConfigSetCommand_instances = new WeakSet(), _CliConfigSetCommand_initTelemetry = function _CliConfigSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        this.telemetryProperties[args.options.key] = args.options.value;
    });
}, _CliConfigSetCommand_initOptions = function _CliConfigSetCommand_initOptions() {
    this.options.unshift({
        option: '-k, --key <key>',
        autocomplete: _a.optionNames
    }, {
        option: '-v, --value <value>'
    });
}, _CliConfigSetCommand_initValidators = function _CliConfigSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (_a.optionNames.indexOf(args.options.key) < 0) {
            return `${args.options.key} is not a valid setting. Allowed values: ${_a.optionNames.join(', ')}`;
        }
        const allowedOutputs = ['text', 'json', 'csv', 'md', 'none'];
        if (args.options.key === settingsNames.output &&
            allowedOutputs.indexOf(args.options.value) === -1) {
            return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${allowedOutputs.join(', ')}`;
        }
        const allowedErrorOutputs = ['stdout', 'stderr'];
        if (args.options.key === settingsNames.errorOutput &&
            allowedErrorOutputs.indexOf(args.options.value) === -1) {
            return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${allowedErrorOutputs.join(', ')}`;
        }
        if (args.options.key === settingsNames.promptListPageSize &&
            typeof args.options.value !== 'number') {
            return `${args.options.value} is not a valid value for the option ${args.options.key}. The value has to be a number.`;
        }
        if (args.options.key === settingsNames.promptListPageSize &&
            args.options.value <= 0) {
            return `${args.options.value} is not a valid value for the option ${args.options.key}. The number has to be higher than 0.`;
        }
        if (args.options.key === settingsNames.helpMode &&
            cli.helpModes.indexOf(args.options.value) === -1) {
            return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${cli.helpModes.join(', ')}`;
        }
        if (args.options.key === settingsNames.authType &&
            !Object.values(AuthType).map(String).includes(args.options.value)) {
            return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${Object.values(AuthType).join(', ')}`;
        }
        if (args.options.key === settingsNames.helpTarget &&
            !cli.helpTargets.includes(args.options.value)) {
            return `${args.options.value} is not a valid value for the option ${args.options.key}. Allowed values: ${cli.helpTargets.join(', ')}`;
        }
        if (args.options.key === settingsNames.clientId &&
            !validation.isValidGuid(args.options.value)) {
            return `${args.options.value} is not a valid value for the option ${args.options.key}. The value has to be a valid GUID.`;
        }
        if (args.options.key === settingsNames.tenantId &&
            !(args.options.value === 'common' || validation.isValidGuid(args.options.value))) {
            return `${args.options.value} is not a valid value for the option ${args.options.key}. The value has to be a valid GUID or 'common'.`;
        }
        return true;
    });
};
CliConfigSetCommand.optionNames = Object.getOwnPropertyNames(settingsNames);
export default new CliConfigSetCommand();
//# sourceMappingURL=config-set.js.map