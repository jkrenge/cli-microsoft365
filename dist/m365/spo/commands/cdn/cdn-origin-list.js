var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCdnOriginListCommand_instances, _SpoCdnOriginListCommand_initTelemetry, _SpoCdnOriginListCommand_initOptions, _SpoCdnOriginListCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCdnOriginListCommand extends SpoCommand {
    get name() {
        return commands.CDN_ORIGIN_LIST;
    }
    get description() {
        return 'List CDN origins settings for the current SharePoint Online tenant';
    }
    constructor() {
        super();
        _SpoCdnOriginListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCdnOriginListCommand_instances, "m", _SpoCdnOriginListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCdnOriginListCommand_instances, "m", _SpoCdnOriginListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCdnOriginListCommand_instances, "m", _SpoCdnOriginListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const cdnTypeString = args.options.type || 'Public';
        const cdnType = cdnTypeString === 'Private' ? 1 : 0;
        try {
            const tenantId = await spo.getTenantId(logger, this.debug);
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving configured origins for ${(cdnType === 1 ? 'Private' : 'Public')} CDN...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnOrigins" Id="22" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="18" Name="${tenantId}" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                const result = json[json.length - 1];
                await logger.log(result);
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoCdnOriginListCommand_instances = new WeakSet(), _SpoCdnOriginListCommand_initTelemetry = function _SpoCdnOriginListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            cdnType: args.options.type || 'Public'
        });
    });
}, _SpoCdnOriginListCommand_initOptions = function _SpoCdnOriginListCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type [type]',
        autocomplete: ['Public', 'Private']
    });
}, _SpoCdnOriginListCommand_initValidators = function _SpoCdnOriginListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type) {
            if (args.options.type !== 'Public' &&
                args.options.type !== 'Private') {
                return `${args.options.type} is not a valid CDN type. Allowed values are Public|Private`;
            }
        }
        return true;
    });
};
export default new SpoCdnOriginListCommand();
//# sourceMappingURL=cdn-origin-list.js.map