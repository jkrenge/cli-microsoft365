var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteRightsRevokeCommand_instances, _SpoHubSiteRightsRevokeCommand_initTelemetry, _SpoHubSiteRightsRevokeCommand_initOptions, _SpoHubSiteRightsRevokeCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHubSiteRightsRevokeCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_RIGHTS_REVOKE;
    }
    get description() {
        return 'Revokes rights to join sites to the specified hub site for one or more principals';
    }
    constructor() {
        super();
        _SpoHubSiteRightsRevokeCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHubSiteRightsRevokeCommand_instances, "m", _SpoHubSiteRightsRevokeCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteRightsRevokeCommand_instances, "m", _SpoHubSiteRightsRevokeCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteRightsRevokeCommand_instances, "m", _SpoHubSiteRightsRevokeCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const revokeRights = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Revoking rights for ${args.options.principals} from ${args.options.hubSiteUrl}...`);
                }
                const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
                const reqDigest = await spo.getRequestDigest(spoAdminUrl);
                const principals = args.options.principals
                    .split(',')
                    .map(p => `<Object Type="String">${formatting.escapeXml(p.trim())}</Object>`)
                    .join('');
                const requestOptions = {
                    url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': reqDigest.FormDigestValue
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="RevokeHubSiteRights" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.hubSiteUrl)}</Parameter><Parameter Type="Array">${principals}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
        };
        if (args.options.force) {
            await revokeRights();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to revoke rights to join sites to the hub site ${args.options.hubSiteUrl} from the specified users?` });
            if (result) {
                await revokeRights();
            }
        }
    }
}
_SpoHubSiteRightsRevokeCommand_instances = new WeakSet(), _SpoHubSiteRightsRevokeCommand_initTelemetry = function _SpoHubSiteRightsRevokeCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoHubSiteRightsRevokeCommand_initOptions = function _SpoHubSiteRightsRevokeCommand_initOptions() {
    this.options.unshift({
        option: '-u, --hubSiteUrl <hubSiteUrl>'
    }, {
        option: '-p, --principals <principals>'
    }, {
        option: '-f, --force'
    });
}, _SpoHubSiteRightsRevokeCommand_initValidators = function _SpoHubSiteRightsRevokeCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.hubSiteUrl));
};
export default new SpoHubSiteRightsRevokeCommand();
//# sourceMappingURL=hubsite-rights-revoke.js.map