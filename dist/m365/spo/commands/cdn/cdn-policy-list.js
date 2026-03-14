var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCdnPolicyListCommand_instances, _SpoCdnPolicyListCommand_initTelemetry, _SpoCdnPolicyListCommand_initOptions, _SpoCdnPolicyListCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCdnPolicyListCommand extends SpoCommand {
    get name() {
        return commands.CDN_POLICY_LIST;
    }
    get description() {
        return 'Lists CDN policies settings for the current SharePoint Online tenant';
    }
    constructor() {
        super();
        _SpoCdnPolicyListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCdnPolicyListCommand_instances, "m", _SpoCdnPolicyListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCdnPolicyListCommand_instances, "m", _SpoCdnPolicyListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCdnPolicyListCommand_instances, "m", _SpoCdnPolicyListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const cdnTypeString = args.options.cdnType || 'Public';
        const cdnType = cdnTypeString === 'Private' ? 1 : 0;
        try {
            const tenantId = await spo.getTenantId(logger, this.debug);
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving configured policies for ${(cdnType === 1 ? 'Private' : 'Public')} CDN...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnPolicies" Id="7" ObjectPathId="3"><Parameters><Parameter Type="Enum">${cdnType}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="3" Name="${tenantId}" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                const result = json[json.length - 1];
                if (this.verbose) {
                    await logger.logToStderr('Configured policies:');
                }
                await logger.log(result.map(o => {
                    const kv = o.split(';');
                    return {
                        Policy: kv[0],
                        Value: kv[1]
                    };
                }));
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoCdnPolicyListCommand_instances = new WeakSet(), _SpoCdnPolicyListCommand_initTelemetry = function _SpoCdnPolicyListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            cdnType: args.options.cdnType || 'Public'
        });
    });
}, _SpoCdnPolicyListCommand_initOptions = function _SpoCdnPolicyListCommand_initOptions() {
    this.options.unshift({
        option: '-t, --cdnType [cdnType]',
        autocomplete: ['Public', 'Private']
    });
}, _SpoCdnPolicyListCommand_initValidators = function _SpoCdnPolicyListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.cdnType) {
            if (args.options.cdnType !== 'Public' &&
                args.options.cdnType !== 'Private') {
                return `${args.options.cdnType} is not a valid CDN type. Allowed values are Public|Private`;
            }
        }
        return true;
    });
};
export default new SpoCdnPolicyListCommand();
//# sourceMappingURL=cdn-policy-list.js.map