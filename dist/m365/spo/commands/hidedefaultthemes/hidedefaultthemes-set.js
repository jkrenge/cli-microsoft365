var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHideDefaultThemesSetCommand_instances, _SpoHideDefaultThemesSetCommand_initTelemetry, _SpoHideDefaultThemesSetCommand_initOptions, _SpoHideDefaultThemesSetCommand_initTypes;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHideDefaultThemesSetCommand extends SpoCommand {
    get name() {
        return commands.HIDEDEFAULTTHEMES_SET;
    }
    get description() {
        return 'Sets the value of the HideDefaultThemes setting';
    }
    constructor() {
        super();
        _SpoHideDefaultThemesSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHideDefaultThemesSetCommand_instances, "m", _SpoHideDefaultThemesSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoHideDefaultThemesSetCommand_instances, "m", _SpoHideDefaultThemesSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHideDefaultThemesSetCommand_instances, "m", _SpoHideDefaultThemesSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Setting the value of the HideDefaultThemes setting to ${args.options.hideDefaultThemes}...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_api/thememanager/SetHideDefaultThemes`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: {
                    hideDefaultThemes: args.options.hideDefaultThemes
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
_SpoHideDefaultThemesSetCommand_instances = new WeakSet(), _SpoHideDefaultThemesSetCommand_initTelemetry = function _SpoHideDefaultThemesSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            hideDefaultThemes: args.options.hideDefaultThemes
        });
    });
}, _SpoHideDefaultThemesSetCommand_initOptions = function _SpoHideDefaultThemesSetCommand_initOptions() {
    this.options.unshift({
        option: '--hideDefaultThemes <hideDefaultThemes>',
        autocomplete: ['true', 'false']
    });
}, _SpoHideDefaultThemesSetCommand_initTypes = function _SpoHideDefaultThemesSetCommand_initTypes() {
    this.types.boolean.push('hideDefaultThemes');
};
export default new SpoHideDefaultThemesSetCommand();
//# sourceMappingURL=hidedefaultthemes-set.js.map