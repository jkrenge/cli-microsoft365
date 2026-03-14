var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoServicePrincipalSetCommand_instances, _SpoServicePrincipalSetCommand_initTelemetry, _SpoServicePrincipalSetCommand_initOptions, _SpoServicePrincipalSetCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoServicePrincipalSetCommand extends SpoCommand {
    get name() {
        return commands.SERVICEPRINCIPAL_SET;
    }
    get description() {
        return 'Enable or disable the service principal';
    }
    constructor() {
        super();
        _SpoServicePrincipalSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalSetCommand_instances, "m", _SpoServicePrincipalSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalSetCommand_instances, "m", _SpoServicePrincipalSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalSetCommand_instances, "m", _SpoServicePrincipalSetCommand_initTypes).call(this);
    }
    alias() {
        return [commands.SP_SET];
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.toggleServicePrincipal(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to ${args.options.enabled ? 'enable' : 'disable'} the service principal?` });
            if (result) {
                await this.toggleServicePrincipal(logger, args);
            }
        }
    }
    async toggleServicePrincipal(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`${(args.options.enabled ? 'Enabling' : 'Disabling')} service principal...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><SetProperty Id="29" ObjectPathId="27" Name="AccountEnabled"><Parameter Type="Boolean">${args.options.enabled}</Parameter></SetProperty><Method Name="Update" Id="30" ObjectPathId="27" /><Query Id="31" ObjectPathId="27"><Query SelectAllProperties="true"><Properties><Property Name="AccountEnabled" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="27" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`
            };
            const response = await request.post(requestOptions);
            const json = JSON.parse(response);
            const responseContent = json[0];
            if (responseContent.ErrorInfo) {
                throw responseContent.ErrorInfo.ErrorMessage;
            }
            const output = json[json.length - 1];
            delete output._ObjectType_;
            await logger.log(output);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoServicePrincipalSetCommand_instances = new WeakSet(), _SpoServicePrincipalSetCommand_initTelemetry = function _SpoServicePrincipalSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            enabled: args.options.enabled
        });
    });
}, _SpoServicePrincipalSetCommand_initOptions = function _SpoServicePrincipalSetCommand_initOptions() {
    this.options.unshift({
        option: '-e, --enabled <enabled>',
        autocomplete: ['true', 'false']
    }, {
        option: '-f, --force'
    });
}, _SpoServicePrincipalSetCommand_initTypes = function _SpoServicePrincipalSetCommand_initTypes() {
    this.types.boolean.push('enabled');
};
export default new SpoServicePrincipalSetCommand();
//# sourceMappingURL=serviceprincipal-set.js.map