var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCdnOriginAddCommand_instances, _SpoCdnOriginAddCommand_initTelemetry, _SpoCdnOriginAddCommand_initOptions, _SpoCdnOriginAddCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCdnOriginAddCommand extends SpoCommand {
    get name() {
        return commands.CDN_ORIGIN_ADD;
    }
    get description() {
        return 'Adds CDN origin to the current SharePoint Online tenant';
    }
    constructor() {
        super();
        _SpoCdnOriginAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoCdnOriginAddCommand_instances, "m", _SpoCdnOriginAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCdnOriginAddCommand_instances, "m", _SpoCdnOriginAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCdnOriginAddCommand_instances, "m", _SpoCdnOriginAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const cdnTypeString = args.options.type || 'Public';
        const cdnType = cdnTypeString === 'Private' ? 1 : 0;
        try {
            const tenantId = await spo.getTenantId(logger, this.debug);
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Adding origin ${args.options.origin} to the ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddTenantCdnOrigin" Id="27" ObjectPathId="23"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.origin)}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="23" Name="${tenantId}" /></ObjectPaths></Request>`
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
_SpoCdnOriginAddCommand_instances = new WeakSet(), _SpoCdnOriginAddCommand_initTelemetry = function _SpoCdnOriginAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            cdnType: args.options.type || 'Public'
        });
    });
}, _SpoCdnOriginAddCommand_initOptions = function _SpoCdnOriginAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type [type]',
        autocomplete: ['Public', 'Private']
    }, {
        option: '-r, --origin <origin>'
    });
}, _SpoCdnOriginAddCommand_initValidators = function _SpoCdnOriginAddCommand_initValidators() {
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
export default new SpoCdnOriginAddCommand();
//# sourceMappingURL=cdn-origin-add.js.map