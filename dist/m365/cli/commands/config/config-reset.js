var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _CliConfigResetCommand_instances, _a, _CliConfigResetCommand_initTelemetry, _CliConfigResetCommand_initOptions, _CliConfigResetCommand_initValidators;
import { cli } from "../../../../cli/cli.js";
import { settingsNames } from "../../../../settingsNames.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";
class CliConfigResetCommand extends AnonymousCommand {
    get name() {
        return commands.CONFIG_RESET;
    }
    get description() {
        return 'Resets the specified CLI configuration option to its default value';
    }
    constructor() {
        super();
        _CliConfigResetCommand_instances.add(this);
        __classPrivateFieldGet(this, _CliConfigResetCommand_instances, "m", _CliConfigResetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _CliConfigResetCommand_instances, "m", _CliConfigResetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _CliConfigResetCommand_instances, "m", _CliConfigResetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.key) {
            cli.getConfig().delete(args.options.key);
        }
        else {
            cli.getConfig().clear();
        }
    }
}
_a = CliConfigResetCommand, _CliConfigResetCommand_instances = new WeakSet(), _CliConfigResetCommand_initTelemetry = function _CliConfigResetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            key: args.options.key
        });
    });
}, _CliConfigResetCommand_initOptions = function _CliConfigResetCommand_initOptions() {
    this.options.unshift({
        option: '-k, --key [key]',
        autocomplete: _a.optionNames
    });
}, _CliConfigResetCommand_initValidators = function _CliConfigResetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.key) {
            if (_a.optionNames.indexOf(args.options.key) < 0) {
                return `${args.options.key} is not a valid setting. Allowed values: ${_a.optionNames.join(', ')}`;
            }
        }
        return true;
    });
};
CliConfigResetCommand.optionNames = Object.getOwnPropertyNames(settingsNames);
export default new CliConfigResetCommand();
//# sourceMappingURL=config-reset.js.map