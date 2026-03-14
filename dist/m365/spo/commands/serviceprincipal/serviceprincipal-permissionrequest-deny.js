var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoServicePrincipalPermissionRequestDenyCommand_instances, _SpoServicePrincipalPermissionRequestDenyCommand_initOptions, _SpoServicePrincipalPermissionRequestDenyCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoServicePrincipalPermissionRequestDenyCommand extends SpoCommand {
    get name() {
        return commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_DENY;
    }
    alias() {
        return [commands.SP_PERMISSIONREQUEST_DENY];
    }
    get description() {
        return 'Denies the specified permission request';
    }
    constructor() {
        super();
        _SpoServicePrincipalPermissionRequestDenyCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalPermissionRequestDenyCommand_instances, "m", _SpoServicePrincipalPermissionRequestDenyCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalPermissionRequestDenyCommand_instances, "m", _SpoServicePrincipalPermissionRequestDenyCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving request digest...`);
            }
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="160" ObjectPathId="159" /><ObjectPath Id="162" ObjectPathId="161" /><ObjectPath Id="164" ObjectPathId="163" /><Method Name="Deny" Id="165" ObjectPathId="163" /></Actions><ObjectPaths><Constructor Id="159" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="161" ParentId="159" Name="PermissionRequests" /><Method Id="163" ParentId="161" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(args.options.id)}}</Parameter></Parameters></Method></ObjectPaths></Request>`
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
_SpoServicePrincipalPermissionRequestDenyCommand_instances = new WeakSet(), _SpoServicePrincipalPermissionRequestDenyCommand_initOptions = function _SpoServicePrincipalPermissionRequestDenyCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
}, _SpoServicePrincipalPermissionRequestDenyCommand_initValidators = function _SpoServicePrincipalPermissionRequestDenyCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new SpoServicePrincipalPermissionRequestDenyCommand();
//# sourceMappingURL=serviceprincipal-permissionrequest-deny.js.map