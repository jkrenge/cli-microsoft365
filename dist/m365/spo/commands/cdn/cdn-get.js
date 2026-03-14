var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCdnGetCommand_instances, _SpoCdnGetCommand_initTelemetry, _SpoCdnGetCommand_initOptions, _SpoCdnGetCommand_initValidators, _SpoCdnGetCommand_initTypes;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCdnGetCommand extends SpoCommand {
    get name() {
        return commands.CDN_GET;
    }
    get description() {
        return 'View current status of the specified Microsoft 365 CDN';
    }
    constructor() {
        super();
        _SpoCdnGetCommand_instances.add(this);
        this.validTypes = ['Public', 'Private'];
        __classPrivateFieldGet(this, _SpoCdnGetCommand_instances, "m", _SpoCdnGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCdnGetCommand_instances, "m", _SpoCdnGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCdnGetCommand_instances, "m", _SpoCdnGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoCdnGetCommand_instances, "m", _SpoCdnGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const cdnTypeString = args.options.type || 'Public';
        const cdnType = cdnTypeString === 'Private' ? 1 : 0;
        try {
            const tenantId = await spo.getTenantId(logger, this.debug);
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving status of ${(cdnType === 1 ? 'Private' : 'Public')} CDN...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="GetTenantCdnEnabled" Id="12" ObjectPathId="8"><Parameters><Parameter Type="Enum">${cdnType}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="8" Name="${tenantId}" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const result = json[json.length - 1];
            if (this.verbose) {
                await logger.logToStderr(`${cdnType === 0 ? 'Public' : 'Private'} CDN at ${spoAdminUrl} is ${result === true ? 'enabled' : 'disabled'}`);
            }
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoCdnGetCommand_instances = new WeakSet(), _SpoCdnGetCommand_initTelemetry = function _SpoCdnGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            cdnType: args.options.type || 'Public'
        });
    });
}, _SpoCdnGetCommand_initOptions = function _SpoCdnGetCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type [type]',
        autocomplete: this.validTypes
    });
}, _SpoCdnGetCommand_initValidators = function _SpoCdnGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type && !this.validTypes.includes(args.options.type)) {
            return `'${args.options.type}' is not a valid CDN type. Allowed values are: ${this.validTypes.join(', ')}.`;
        }
        return true;
    });
}, _SpoCdnGetCommand_initTypes = function _SpoCdnGetCommand_initTypes() {
    this.types.string.push('type');
};
export default new SpoCdnGetCommand();
//# sourceMappingURL=cdn-get.js.map