var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoServicePrincipalGrantRevokeCommand_instances, _SpoServicePrincipalGrantRevokeCommand_initOptions;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoServicePrincipalGrantRevokeCommand extends SpoCommand {
    get name() {
        return commands.SERVICEPRINCIPAL_GRANT_REVOKE;
    }
    get description() {
        return 'Revokes the specified set of permissions granted to the service principal';
    }
    alias() {
        return [commands.SP_GRANT_REVOKE];
    }
    constructor() {
        super();
        _SpoServicePrincipalGrantRevokeCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalGrantRevokeCommand_instances, "m", _SpoServicePrincipalGrantRevokeCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving request digest...`);
            }
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (args.options.scope) {
                // revoke a single scope
                // #1 get the grant
                const getGrantRequestOptions = {
                    url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': reqDigest.FormDigestValue
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="3" ParentId="1" Name="PermissionGrants" /><Method Id="5" ParentId="3" Name="GetByObjectId"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method></ObjectPaths></Request>`
                };
                const grantRequestRes = await request.post(getGrantRequestOptions);
                const grantRequestJson = JSON.parse(grantRequestRes);
                const responseInfo = grantRequestJson[0];
                if (responseInfo.ErrorInfo) {
                    throw responseInfo.ErrorInfo.ErrorMessage;
                }
                const grantInfo = grantRequestJson.pop();
                // #2 remove the scope from the grant
                const removeScopeRequestOptions = {
                    url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': reqDigest.FormDigestValue
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="Remove" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">${grantInfo.ClientId}</Parameter><Parameter Type="String">Microsoft Graph</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.scope)}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="8" ParentId="1" Name="GrantManager" /><Constructor Id="1" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`
                };
                const removeScopeRes = await request.post(removeScopeRequestOptions);
                const removeScopeResJson = JSON.parse(removeScopeRes);
                const removeScopeResponse = removeScopeResJson[0];
                if (removeScopeResponse.ErrorInfo) {
                    throw removeScopeResponse.ErrorInfo.ErrorMessage;
                }
            }
            else {
                // revoke the whole grant
                const requestOptions = {
                    url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                    headers: {
                        'X-RequestDigest': reqDigest.FormDigestValue
                    },
                    data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Method Name="DeleteObject" Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionGrants" /><Method Id="13" ParentId="11" Name="GetByObjectId"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method></ObjectPaths></Request>`
                };
                const res = await request.post(requestOptions);
                const json = JSON.parse(res);
                const response = json[0];
                if (response.ErrorInfo) {
                    throw response.ErrorInfo.ErrorMessage;
                }
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoServicePrincipalGrantRevokeCommand_instances = new WeakSet(), _SpoServicePrincipalGrantRevokeCommand_initOptions = function _SpoServicePrincipalGrantRevokeCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-s, --scope [scope]'
    });
};
export default new SpoServicePrincipalGrantRevokeCommand();
//# sourceMappingURL=serviceprincipal-grant-revoke.js.map