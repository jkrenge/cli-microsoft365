var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoThemeSetCommand_instances, _SpoThemeSetCommand_initTelemetry, _SpoThemeSetCommand_initOptions, _SpoThemeSetCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoThemeSetCommand extends SpoCommand {
    get name() {
        return commands.THEME_SET;
    }
    get description() {
        return 'Add or update a theme';
    }
    constructor() {
        super();
        _SpoThemeSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoThemeSetCommand_instances, "m", _SpoThemeSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoThemeSetCommand_instances, "m", _SpoThemeSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoThemeSetCommand_instances, "m", _SpoThemeSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoAdminUrl);
            const palette = JSON.parse(args.options.theme);
            if (this.debug) {
                await logger.logToStderr('');
                await logger.logToStderr('Palette');
                await logger.logToStderr(JSON.stringify(palette));
            }
            const isInverted = args.options.isInverted ? true : false;
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="UpdateTenantTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter><Parameter Type="String">{"isInverted":${isInverted},"name":"${formatting.escapeXml(args.options.name)}","palette":${JSON.stringify(palette)}}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const contents = json.find(x => { return x['ErrorInfo']; });
            if (contents && contents.ErrorInfo) {
                throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoThemeSetCommand_instances = new WeakSet(), _SpoThemeSetCommand_initTelemetry = function _SpoThemeSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            inverted: (!(!args.options.isInverted)).toString()
        });
    });
}, _SpoThemeSetCommand_initOptions = function _SpoThemeSetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-t, --theme <theme>'
    }, {
        option: '--isInverted'
    });
}, _SpoThemeSetCommand_initValidators = function _SpoThemeSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidTheme(args.options.theme)) {
            return 'The specified theme is not valid';
        }
        return true;
    });
};
export default new SpoThemeSetCommand();
//# sourceMappingURL=theme-set.js.map