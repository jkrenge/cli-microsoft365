var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoServicePrincipalPermissionRequestApproveCommand_instances, _SpoServicePrincipalPermissionRequestApproveCommand_initOptions, _SpoServicePrincipalPermissionRequestApproveCommand_initValidators, _SpoServicePrincipalPermissionRequestApproveCommand_initTelemetry, _SpoServicePrincipalPermissionRequestApproveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoServicePrincipalPermissionRequestListCommand from './serviceprincipal-permissionrequest-list.js';
class SpoServicePrincipalPermissionRequestApproveCommand extends SpoCommand {
    get name() {
        return commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_APPROVE;
    }
    get description() {
        return 'Approves the specified permission request';
    }
    alias() {
        return [commands.SP_PERMISSIONREQUEST_APPROVE];
    }
    constructor() {
        super();
        _SpoServicePrincipalPermissionRequestApproveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalPermissionRequestApproveCommand_instances, "m", _SpoServicePrincipalPermissionRequestApproveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalPermissionRequestApproveCommand_instances, "m", _SpoServicePrincipalPermissionRequestApproveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalPermissionRequestApproveCommand_instances, "m", _SpoServicePrincipalPermissionRequestApproveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoServicePrincipalPermissionRequestApproveCommand_instances, "m", _SpoServicePrincipalPermissionRequestApproveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving request digest...`);
            }
            const permissionRequestIds = await this.getAllPendingPermissionRequests(args);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            const response = [];
            await permissionRequestIds.reduce(async (previousPromise, nextPermissionRequestId) => {
                return previousPromise.then(() => {
                    return this.approvePermissionRequest(nextPermissionRequestId, reqDigest, spoAdminUrl).then(result => response.push(result));
                });
            }, Promise.resolve());
            await logger.log(response.length === 1 ? response[0] : response);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    async getAllPendingPermissionRequests(args) {
        if (args.options.id) {
            return [args.options.id];
        }
        else {
            const options = {
                debug: this.debug,
                verbose: this.verbose
            };
            const output = await cli.executeCommandWithOutput(spoServicePrincipalPermissionRequestListCommand, { options: { ...options, _: [] } });
            const getPermissionRequestsOutput = JSON.parse(output.stdout);
            if (args.options.resource) {
                return getPermissionRequestsOutput.filter((x) => x.Resource === args.options.resource).map((x) => { return x.Id; });
            }
            return getPermissionRequestsOutput.map((x) => { return x.Id; });
        }
    }
    async approvePermissionRequest(permissionRequestId, reqDigest, spoAdminUrl) {
        const requestOptions = {
            url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
                'X-RequestDigest': reqDigest.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${formatting.escapeXml(permissionRequestId)}}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>`
        };
        const res = await request.post(requestOptions);
        const json = JSON.parse(res);
        const response = json[0];
        if (response.ErrorInfo) {
            throw response.ErrorInfo.ErrorMessage;
        }
        else {
            const output = json[json.length - 1];
            delete output._ObjectType_;
            return output;
        }
    }
}
_SpoServicePrincipalPermissionRequestApproveCommand_instances = new WeakSet(), _SpoServicePrincipalPermissionRequestApproveCommand_initOptions = function _SpoServicePrincipalPermissionRequestApproveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '--all'
    }, {
        option: '--resource [resource]'
    });
}, _SpoServicePrincipalPermissionRequestApproveCommand_initValidators = function _SpoServicePrincipalPermissionRequestApproveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _SpoServicePrincipalPermissionRequestApproveCommand_initTelemetry = function _SpoServicePrincipalPermissionRequestApproveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            resource: typeof args.options.resource !== 'undefined',
            all: !!args.options.all
        });
    });
}, _SpoServicePrincipalPermissionRequestApproveCommand_initOptionSets = function _SpoServicePrincipalPermissionRequestApproveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'all', 'resource'] });
};
export default new SpoServicePrincipalPermissionRequestApproveCommand();
//# sourceMappingURL=serviceprincipal-permissionrequest-approve.js.map