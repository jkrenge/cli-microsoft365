var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoThemeRemoveCommand_instances, _SpoThemeRemoveCommand_initTelemetry, _SpoThemeRemoveCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoThemeRemoveCommand extends SpoCommand {
    get name() {
        return commands.THEME_REMOVE;
    }
    get description() {
        return 'Removes existing theme';
    }
    constructor() {
        super();
        _SpoThemeRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoThemeRemoveCommand_instances, "m", _SpoThemeRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoThemeRemoveCommand_instances, "m", _SpoThemeRemoveCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeTheme(logger, args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the theme` });
            if (result) {
                await this.removeTheme(logger, args.options);
            }
        }
    }
    async removeTheme(logger, options) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Removing theme from tenant...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_api/thememanager/DeleteTenantTheme`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: {
                    name: options.name
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoThemeRemoveCommand_instances = new WeakSet(), _SpoThemeRemoveCommand_initTelemetry = function _SpoThemeRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoThemeRemoveCommand_initOptions = function _SpoThemeRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-f, --force'
    });
};
export default new SpoThemeRemoveCommand();
//# sourceMappingURL=theme-remove.js.map