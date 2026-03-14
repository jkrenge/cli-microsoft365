var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoThemeGetCommand_instances, _SpoThemeGetCommand_initOptions;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoThemeGetCommand extends SpoCommand {
    get name() {
        return commands.THEME_GET;
    }
    get description() {
        return 'Gets custom theme information';
    }
    constructor() {
        super();
        _SpoThemeGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoThemeGetCommand_instances, "m", _SpoThemeGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Getting ${args.options.name} theme from tenant...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Query Id="15" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="13" ParentId="11" Name="GetTenantTheme"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.name)}</Parameter></Parameters></Method></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const contents = json.find(x => { return x['ErrorInfo']; });
            if (contents && contents.ErrorInfo) {
                throw contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error';
            }
            const theme = json[6];
            delete theme._ObjectType_;
            await logger.log(theme);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoThemeGetCommand_instances = new WeakSet(), _SpoThemeGetCommand_initOptions = function _SpoThemeGetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    });
};
export default new SpoThemeGetCommand();
//# sourceMappingURL=theme-get.js.map