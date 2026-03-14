var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoServicePrincipalGrantAddCommand_instances, _SpoServicePrincipalGrantAddCommand_initOptions;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoServicePrincipalGrantAddCommand extends SpoCommand {
    get name() {
        return commands.SERVICEPRINCIPAL_GRANT_ADD;
    }
    get description() {
        return 'Grants the service principal permission to the specified API';
    }
    alias() {
        return [commands.SP_GRANT_ADD];
    }
    constructor() {
        super();
        _SpoServicePrincipalGrantAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalGrantAddCommand_instances, "m", _SpoServicePrincipalGrantAddCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving request digest...`);
            }
            const resDigest = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': resDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><ObjectPath Id="8" ObjectPathId="7" /><Query Id="9" ObjectPathId="7"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="5" ParentId="3" Name="PermissionRequests" /><Method Id="7" ParentId="5" Name="Approve"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.resource)}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.scope)}</Parameter></Parameters></Method></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                const result = json[json.length - 1];
                delete result._ObjectType_;
                await logger.log(result);
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoServicePrincipalGrantAddCommand_instances = new WeakSet(), _SpoServicePrincipalGrantAddCommand_initOptions = function _SpoServicePrincipalGrantAddCommand_initOptions() {
    this.options.unshift({
        option: '-r, --resource <resource>'
    }, {
        option: '-s, --scope <scope>'
    });
};
export default new SpoServicePrincipalGrantAddCommand();
//# sourceMappingURL=serviceprincipal-grant-add.js.map