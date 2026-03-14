var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoCdnSetCommand_instances, _SpoCdnSetCommand_initTelemetry, _SpoCdnSetCommand_initOptions, _SpoCdnSetCommand_initTypes, _SpoCdnSetCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoCdnSetCommand extends SpoCommand {
    get name() {
        return commands.CDN_SET;
    }
    get description() {
        return 'Enable or disable the specified Microsoft 365 CDN';
    }
    constructor() {
        super();
        _SpoCdnSetCommand_instances.add(this);
        this.validTypes = ['Public', 'Private', 'Both'];
        __classPrivateFieldGet(this, _SpoCdnSetCommand_instances, "m", _SpoCdnSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoCdnSetCommand_instances, "m", _SpoCdnSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoCdnSetCommand_instances, "m", _SpoCdnSetCommand_initTypes).call(this);
        __classPrivateFieldGet(this, _SpoCdnSetCommand_instances, "m", _SpoCdnSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const cdnTypeString = args.options.type || 'Public';
        const enabled = args.options.enabled;
        let cdnType;
        switch (cdnTypeString) {
            case "Private": {
                cdnType = 1;
                break;
            }
            case "Both": {
                cdnType = 2;
                break;
            }
            default: {
                cdnType = 0;
                break;
            }
        }
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            let requestBody = '';
            if (cdnType === 2) {
                if (args.options.noDefaultOrigins) {
                    if (this.verbose) {
                        await logger.logToStderr(`${(enabled ? 'Enabling' : 'Disabling')} Public and Private CDNs without default origins. Please wait, this might take a moment...`);
                    }
                    requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><Method Name="SetTenantCdnEnabled" Id="13" ObjectPathId="11"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="14" ObjectPathId="11"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
                }
                else {
                    if (this.verbose) {
                        await logger.logToStderr(`${(enabled ? 'Enabling' : 'Disabling')} Public and Private CDNs with default origins. Please wait, this might take a moment...`);
                    }
                    requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="96" ObjectPathId="95" /><Method Name="SetTenantCdnEnabled" Id="97" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="SetTenantCdnEnabled" Id="98" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method>${(enabled ? '<Method Name="CreateTenantCdnDefaultOrigins" Id="99" ObjectPathId="95"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="100" ObjectPathId="95"><Parameters><Parameter Type="Enum">0</Parameter></Parameters></Method>' : '')}</Actions><ObjectPaths><Constructor Id="95" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
                }
            }
            else {
                if (args.options.noDefaultOrigins) {
                    if (this.verbose) {
                        await logger.logToStderr(`${(enabled ? 'Enabling' : 'Disabling')} ${(cdnType === 1 ? 'Private' : 'Public')} CDN without default origins. Please wait, this might take a moment...`);
                    }
                    requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
                }
                else {
                    if (this.verbose) {
                        await logger.logToStderr(`${(enabled ? 'Enabling' : 'Disabling')} ${(cdnType === 1 ? 'Private' : 'Public')} CDN. Please wait, this might take a moment...`);
                    }
                    if (enabled) {
                        requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method><Method Name="CreateTenantCdnDefaultOrigins" Id="21" ObjectPathId="18"><Parameters><Parameter Type="Enum">${(cdnType === 1 ? 1 : 0)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
                    }
                    else {
                        requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="19" ObjectPathId="18" /><Method Name="SetTenantCdnEnabled" Id="20" ObjectPathId="18"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="Boolean">${enabled}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="18" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`;
                    }
                }
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: requestBody
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
_SpoCdnSetCommand_instances = new WeakSet(), _SpoCdnSetCommand_initTelemetry = function _SpoCdnSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            cdnType: args.options.type || 'Public',
            enabled: args.options.enabled,
            noDefaultOrigins: !!args.options.noDefaultOrigins
        });
    });
}, _SpoCdnSetCommand_initOptions = function _SpoCdnSetCommand_initOptions() {
    this.options.unshift({
        option: '-e, --enabled <enabled>',
        autocomplete: ['true', 'false']
    }, {
        option: '-t, --type [type]',
        autocomplete: this.validTypes
    }, {
        option: '--noDefaultOrigins'
    });
}, _SpoCdnSetCommand_initTypes = function _SpoCdnSetCommand_initTypes() {
    this.types.boolean.push('enabled', 'noDefaultOrigins');
    this.types.string.push('type');
}, _SpoCdnSetCommand_initValidators = function _SpoCdnSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type) {
            if (args.options.type !== 'Public' && args.options.type !== 'Both' &&
                args.options.type !== 'Private') {
                return `${args.options.type} is not a valid CDN type. Allowed values are ${this.validTypes.join(', ')}.`;
            }
        }
        return true;
    });
};
export default new SpoCdnSetCommand();
//# sourceMappingURL=cdn-set.js.map