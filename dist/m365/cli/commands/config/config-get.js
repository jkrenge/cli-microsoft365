var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _CliConfigGetCommand_instances, _a, _CliConfigGetCommand_initTelemetry, _CliConfigGetCommand_initOptions, _CliConfigGetCommand_initValidators;
import { cli } from "../../../../cli/cli.js";
import { settingsNames } from "../../../../settingsNames.js";
import AnonymousCommand from "../../../base/AnonymousCommand.js";
import commands from "../../commands.js";
class CliConfigGetCommand extends AnonymousCommand {
    get name() {
        return commands.CONFIG_GET;
    }
    get description() {
        return 'Gets value of a CLI for Microsoft 365 configuration option';
    }
    constructor() {
        super();
        _CliConfigGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _CliConfigGetCommand_instances, "m", _CliConfigGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _CliConfigGetCommand_instances, "m", _CliConfigGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _CliConfigGetCommand_instances, "m", _CliConfigGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        await logger.log(cli.getConfig().get(args.options.key));
    }
}
_a = CliConfigGetCommand, _CliConfigGetCommand_instances = new WeakSet(), _CliConfigGetCommand_initTelemetry = function _CliConfigGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            key: args.options.key
        });
    });
}, _CliConfigGetCommand_initOptions = function _CliConfigGetCommand_initOptions() {
    this.options.unshift({
        option: '-k, --key <key>',
        autocomplete: _a.optionNames
    });
}, _CliConfigGetCommand_initValidators = function _CliConfigGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (_a.optionNames.indexOf(args.options.key) < 0) {
            return `${args.options.key} is not a valid setting. Allowed values: ${_a.optionNames.join(', ')}`;
        }
        return true;
    });
};
CliConfigGetCommand.optionNames = Object.getOwnPropertyNames(settingsNames);
export default new CliConfigGetCommand();
//# sourceMappingURL=config-get.js.map