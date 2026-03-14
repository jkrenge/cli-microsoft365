var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoStorageEntityRemoveCommand_instances, _SpoStorageEntityRemoveCommand_initTelemetry, _SpoStorageEntityRemoveCommand_initOptions, _SpoStorageEntityRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoStorageEntityRemoveCommand extends SpoCommand {
    get name() {
        return commands.STORAGEENTITY_REMOVE;
    }
    get description() {
        return 'Removes tenant property stored on the specified SharePoint Online app catalog';
    }
    constructor() {
        super();
        _SpoStorageEntityRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoStorageEntityRemoveCommand_instances, "m", _SpoStorageEntityRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoStorageEntityRemoveCommand_instances, "m", _SpoStorageEntityRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoStorageEntityRemoveCommand_instances, "m", _SpoStorageEntityRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeTenantProperty(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to delete the ${args.options.key} tenant property?` });
            if (result) {
                await this.removeTenantProperty(logger, args);
            }
        }
    }
    async removeTenantProperty(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing tenant property ${args.options.key} from ${args.options.appCatalogUrl}...`);
        }
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const digestInfo = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': digestInfo.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.key)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.appCatalogUrl)}</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoStorageEntityRemoveCommand_instances = new WeakSet(), _SpoStorageEntityRemoveCommand_initTelemetry = function _SpoStorageEntityRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoStorageEntityRemoveCommand_initOptions = function _SpoStorageEntityRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --appCatalogUrl <appCatalogUrl>'
    }, {
        option: '-k, --key <key>'
    }, {
        option: '-f, --force'
    });
}, _SpoStorageEntityRemoveCommand_initValidators = function _SpoStorageEntityRemoveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.appCatalogUrl));
};
export default new SpoStorageEntityRemoveCommand();
//# sourceMappingURL=storageentity-remove.js.map