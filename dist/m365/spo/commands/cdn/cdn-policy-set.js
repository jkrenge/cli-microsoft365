var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCdnPolicySetCommand_instances, _SpoCdnPolicySetCommand_initTelemetry, _SpoCdnPolicySetCommand_initOptions, _SpoCdnPolicySetCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCdnPolicySetCommand extends SpoCommand {
    get name() {
        return commands.CDN_POLICY_SET;
    }
    get description() {
        return 'Sets CDN policy value for the current SharePoint Online tenant';
    }
    constructor() {
        super();
        _SpoCdnPolicySetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCdnPolicySetCommand_instances, "m", _SpoCdnPolicySetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCdnPolicySetCommand_instances, "m", _SpoCdnPolicySetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCdnPolicySetCommand_instances, "m", _SpoCdnPolicySetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const cdnTypeString = args.options.cdnType || 'Public';
        const cdnType = cdnTypeString === 'Private' ? 1 : 0;
        try {
            const tenantId = await spo.getTenantId(logger, this.debug);
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Configuring policy on the ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
            }
            let policyId = -1;
            switch (args.options.policy) {
                case "IncludeFileExtensions":
                    policyId = 0;
                    break;
                case "ExcludeRestrictedSiteClassifications":
                    policyId = 1;
                    break;
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="SetTenantCdnPolicy" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Enum">${policyId}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.value)}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="${tenantId}" /></ObjectPaths></Request>`
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
_SpoCdnPolicySetCommand_instances = new WeakSet(), _SpoCdnPolicySetCommand_initTelemetry = function _SpoCdnPolicySetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            cdnType: args.options.cdnType || 'Public',
            policy: args.options.policy
        });
    });
}, _SpoCdnPolicySetCommand_initOptions = function _SpoCdnPolicySetCommand_initOptions() {
    this.options.unshift({
        option: '-t, --cdnType [cdnType]',
        autocomplete: ['Public', 'Private']
    }, {
        option: '-p, --policy <policy>',
        autocomplete: ['IncludeFileExtensions', 'ExcludeRestrictedSiteClassifications']
    }, {
        option: '-v, --value <value>'
    });
}, _SpoCdnPolicySetCommand_initValidators = function _SpoCdnPolicySetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.cdnType) {
            if (args.options.cdnType !== 'Public' &&
                args.options.cdnType !== 'Private') {
                return `${args.options.cdnType} is not a valid CDN type. Allowed values are Public|Private`;
            }
        }
        if (!args.options.policy ||
            (args.options.policy !== 'IncludeFileExtensions' &&
                args.options.policy !== 'ExcludeRestrictedSiteClassifications')) {
            return `${args.options.policy} is not a valid CDN policy. Allowed values are IncludeFileExtensions|ExcludeRestrictedSiteClassifications`;
        }
        return true;
    });
};
export default new SpoCdnPolicySetCommand();
//# sourceMappingURL=cdn-policy-set.js.map