var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAppCatalogAddCommand_instances, _SpoSiteAppCatalogAddCommand_initOptions, _SpoSiteAppCatalogAddCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteAppCatalogAddCommand extends SpoCommand {
    get name() {
        return commands.SITE_APPCATALOG_ADD;
    }
    get description() {
        return 'Creates a site collection app catalog in the specified site';
    }
    constructor() {
        super();
        _SpoSiteAppCatalogAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteAppCatalogAddCommand_instances, "m", _SpoSiteAppCatalogAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppCatalogAddCommand_instances, "m", _SpoSiteAppCatalogAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const url = args.options.siteUrl;
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const requestDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Adding site collection app catalog...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': requestDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="38" ObjectPathId="37" /><ObjectPath Id="40" ObjectPathId="39" /><ObjectPath Id="42" ObjectPathId="41" /><ObjectPath Id="44" ObjectPathId="43" /><ObjectPath Id="46" ObjectPathId="45" /><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Constructor Id="37" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="39" ParentId="37" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Property Id="41" ParentId="39" Name="RootWeb" /><Property Id="43" ParentId="41" Name="TenantAppCatalog" /><Property Id="45" ParentId="43" Name="SiteCollectionAppCatalogsSites" /><Method Id="47" ParentId="45" Name="Add"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr('Site collection app catalog created');
                }
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoSiteAppCatalogAddCommand_instances = new WeakSet(), _SpoSiteAppCatalogAddCommand_initOptions = function _SpoSiteAppCatalogAddCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    });
}, _SpoSiteAppCatalogAddCommand_initValidators = function _SpoSiteAppCatalogAddCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.siteUrl));
};
export default new SpoSiteAppCatalogAddCommand();
//# sourceMappingURL=site-appcatalog-add.js.map