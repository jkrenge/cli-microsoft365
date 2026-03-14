var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteAppCatalogRemoveCommand_instances, _SpoSiteAppCatalogRemoveCommand_initTelemetry, _SpoSiteAppCatalogRemoveCommand_initOptions, _SpoSiteAppCatalogRemoveCommand_initValidators, _SpoSiteAppCatalogRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteAppCatalogRemoveCommand extends SpoCommand {
    get name() {
        return commands.SITE_APPCATALOG_REMOVE;
    }
    get description() {
        return 'Removes site collection app catalog from the specified site';
    }
    constructor() {
        super();
        _SpoSiteAppCatalogRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteAppCatalogRemoveCommand_instances, "m", _SpoSiteAppCatalogRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppCatalogRemoveCommand_instances, "m", _SpoSiteAppCatalogRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppCatalogRemoveCommand_instances, "m", _SpoSiteAppCatalogRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteAppCatalogRemoveCommand_instances, "m", _SpoSiteAppCatalogRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeSiteAppcatalog = async () => {
            const url = args.options.siteUrl;
            if (this.verbose) {
                await logger.logToStderr(`Disabling site collection app catalog...`);
            }
            try {
                const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
                const requestDigest = await spo.getRequestDigest(spoAdminUrl);
                const requestOptions = {
                    url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': requestDigest.FormDigestValue
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="50" ObjectPathId="49" /><ObjectPath Id="52" ObjectPathId="51" /><ObjectPath Id="54" ObjectPathId="53" /><ObjectPath Id="56" ObjectPathId="55" /><ObjectPath Id="58" ObjectPathId="57" /><Method Name="Remove" Id="59" ObjectPathId="57"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="49" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="51" ParentId="49" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method><Property Id="53" ParentId="51" Name="RootWeb" /><Property Id="55" ParentId="53" Name="TenantAppCatalog" /><Property Id="57" ParentId="55" Name="SiteCollectionAppCatalogsSites" /></ObjectPaths></Request>`
                };
                const res = await request.post(requestOptions);
                const json = JSON.parse(res);
                const response = json[0];
                if (response.ErrorInfo) {
                    throw response.ErrorInfo.ErrorMessage;
                }
                else {
                    if (this.verbose) {
                        await logger.logToStderr('Site collection app catalog disabled');
                    }
                }
            }
            catch (err) {
                this.handleRejectedPromise(err);
            }
        };
        if (args.options.force) {
            await removeSiteAppcatalog();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app catalog from ${args.options.siteUrl}?` });
            if (result) {
                await removeSiteAppcatalog();
            }
        }
    }
}
_SpoSiteAppCatalogRemoveCommand_instances = new WeakSet(), _SpoSiteAppCatalogRemoveCommand_initTelemetry = function _SpoSiteAppCatalogRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _SpoSiteAppCatalogRemoveCommand_initOptions = function _SpoSiteAppCatalogRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteAppCatalogRemoveCommand_initValidators = function _SpoSiteAppCatalogRemoveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.siteUrl));
}, _SpoSiteAppCatalogRemoveCommand_initTypes = function _SpoSiteAppCatalogRemoveCommand_initTypes() {
    this.types.string.push('siteUrl');
};
export default new SpoSiteAppCatalogRemoveCommand();
//# sourceMappingURL=site-appcatalog-remove.js.map