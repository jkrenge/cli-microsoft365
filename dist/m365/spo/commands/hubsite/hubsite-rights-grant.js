var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteRightsGrantCommand_instances, _SpoHubSiteRightsGrantCommand_initOptions, _SpoHubSiteRightsGrantCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHubSiteRightsGrantCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_RIGHTS_GRANT;
    }
    get description() {
        return 'Grants permissions to join the hub site for one or more principals';
    }
    constructor() {
        super();
        _SpoHubSiteRightsGrantCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHubSiteRightsGrantCommand_instances, "m", _SpoHubSiteRightsGrantCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteRightsGrantCommand_instances, "m", _SpoHubSiteRightsGrantCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Granting permissions to join the hub site ${args.options.hubSiteUrl} to principals ${args.options.principals}...`);
            }
            const principals = args.options.principals
                .split(',')
                .map(p => `<Object Type="String">${formatting.escapeXml(p.trim())}</Object>`)
                .join('');
            const grantedRights = '1';
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="37" ObjectPathId="36" /><Method Name="GrantHubSiteRights" Id="38" ObjectPathId="36"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.hubSiteUrl)}</Parameter><Parameter Type="Array">${principals}</Parameter><Parameter Type="Enum">${grantedRights}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="36" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoHubSiteRightsGrantCommand_instances = new WeakSet(), _SpoHubSiteRightsGrantCommand_initOptions = function _SpoHubSiteRightsGrantCommand_initOptions() {
    this.options.unshift({
        option: '-u, --hubSiteUrl <hubSiteUrl>'
    }, {
        option: '-p, --principals <principals>'
    }, {
        option: '-r, --rights <rights>',
        autocomplete: ['Join']
    });
}, _SpoHubSiteRightsGrantCommand_initValidators = function _SpoHubSiteRightsGrantCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.hubSiteUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.rights !== 'Join') {
            return `${args.options.rights} is not a valid rights value. Allowed values Join`;
        }
        return true;
    });
};
export default new SpoHubSiteRightsGrantCommand();
//# sourceMappingURL=hubsite-rights-grant.js.map