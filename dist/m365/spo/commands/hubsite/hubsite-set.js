var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteSetCommand_instances, _SpoHubSiteSetCommand_initTelemetry, _SpoHubSiteSetCommand_initOptions, _SpoHubSiteSetCommand_initValidators, _SpoHubSiteSetCommand_initTypes;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHubSiteSetCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_SET;
    }
    get description() {
        return 'Updates properties of the specified hub site';
    }
    constructor() {
        super();
        _SpoHubSiteSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHubSiteSetCommand_instances, "m", _SpoHubSiteSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteSetCommand_instances, "m", _SpoHubSiteSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteSetCommand_instances, "m", _SpoHubSiteSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteSetCommand_instances, "m", _SpoHubSiteSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Updating hub site ${args.options.id}...`);
            }
            const title = typeof args.options.title === 'string' ? `<SetProperty Id="13" ObjectPathId="10" Name="Title"><Parameter Type="String">${formatting.escapeXml(args.options.title)}</Parameter></SetProperty>` : '';
            const description = typeof args.options.description === 'string' ? `<SetProperty Id="15" ObjectPathId="10" Name="Description"><Parameter Type="String">${formatting.escapeXml(args.options.description)}</Parameter></SetProperty>` : '';
            const logoUrl = typeof args.options.logoUrl === 'string' ? `<SetProperty Id="14" ObjectPathId="10" Name="LogoUrl"><Parameter Type="String">${formatting.escapeXml(args.options.logoUrl)}</Parameter></SetProperty>` : '';
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><ObjectPath Id="11" ObjectPathId="10" /><Query Id="12" ObjectPathId="10"><Query SelectAllProperties="true"><Properties /></Query></Query>${title}${logoUrl}${description}<Method Name="Update" Id="16" ObjectPathId="10" /></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="10" ParentId="8" Name="GetHubSitePropertiesById"><Parameters><Parameter Type="Guid">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                const hubSite = json.pop();
                delete hubSite._ObjectType_;
                hubSite.ID = hubSite.ID.replace('/Guid(', '').replace(')/', '');
                hubSite.SiteId = hubSite.SiteId.replace('/Guid(', '').replace(')/', '');
                await logger.log(hubSite);
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoHubSiteSetCommand_instances = new WeakSet(), _SpoHubSiteSetCommand_initTelemetry = function _SpoHubSiteSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            title: typeof args.options.title === 'string',
            description: typeof args.options.description === 'string',
            logoUrl: typeof args.options.logoUrl === 'string'
        });
    });
}, _SpoHubSiteSetCommand_initOptions = function _SpoHubSiteSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-l, --logoUrl [logoUrl]'
    });
}, _SpoHubSiteSetCommand_initValidators = function _SpoHubSiteSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (!args.options.title &&
            !args.options.description &&
            !args.options.logoUrl) {
            return 'Specify title, description or logoUrl to update';
        }
        return true;
    });
}, _SpoHubSiteSetCommand_initTypes = function _SpoHubSiteSetCommand_initTypes() {
    this.types.string.push('t', 'title', 'd', 'description', 'l', 'logoUrl');
};
export default new SpoHubSiteSetCommand();
//# sourceMappingURL=hubsite-set.js.map