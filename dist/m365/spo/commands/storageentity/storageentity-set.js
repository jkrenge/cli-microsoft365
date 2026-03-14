var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoStorageEntitySetCommand_instances, _SpoStorageEntitySetCommand_initTelemetry, _SpoStorageEntitySetCommand_initOptions, _SpoStorageEntitySetCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoStorageEntitySetCommand extends SpoCommand {
    get name() {
        return commands.STORAGEENTITY_SET;
    }
    get description() {
        return 'Sets tenant property on the specified SharePoint Online app catalog';
    }
    constructor() {
        super();
        _SpoStorageEntitySetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoStorageEntitySetCommand_instances, "m", _SpoStorageEntitySetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoStorageEntitySetCommand_instances, "m", _SpoStorageEntitySetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoStorageEntitySetCommand_instances, "m", _SpoStorageEntitySetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoAdminUrl);
            if (this.verbose) {
                await logger.logToStderr(`Setting tenant property ${args.options.key} in ${args.options.appCatalogUrl}...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.key)}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.value)}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.description || '')}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.comment || '')}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.appCatalogUrl)}</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                if (this.verbose && response.ErrorInfo.ErrorMessage.indexOf('Access denied.') > -1) {
                    await logger.logToStderr('');
                    await logger.logToStderr(`This error is often caused by invalid URL of the app catalog site. Verify, that the URL you specified as an argument of the ${commands.STORAGEENTITY_SET} command is a valid app catalog URL and try again.`);
                    await logger.logToStderr('');
                }
                throw response.ErrorInfo.ErrorMessage;
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoStorageEntitySetCommand_instances = new WeakSet(), _SpoStorageEntitySetCommand_initTelemetry = function _SpoStorageEntitySetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: (!(!args.options.description)).toString(),
            comment: (!(!args.options.comment)).toString()
        });
    });
}, _SpoStorageEntitySetCommand_initOptions = function _SpoStorageEntitySetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --appCatalogUrl <appCatalogUrl>'
    }, {
        option: '-k, --key <key>'
    }, {
        option: '-v, --value <value>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '-c, --comment [comment]'
    });
}, _SpoStorageEntitySetCommand_initValidators = function _SpoStorageEntitySetCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.appCatalogUrl));
};
export default new SpoStorageEntitySetCommand();
//# sourceMappingURL=storageentity-set.js.map