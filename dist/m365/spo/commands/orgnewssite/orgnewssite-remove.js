var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoOrgNewsSiteRemoveCommand_instances, _SpoOrgNewsSiteRemoveCommand_initTelemetry, _SpoOrgNewsSiteRemoveCommand_initOptions, _SpoOrgNewsSiteRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoOrgNewsSiteRemoveCommand extends SpoCommand {
    get name() {
        return commands.ORGNEWSSITE_REMOVE;
    }
    get description() {
        return 'Removes a site from the list of organizational news sites';
    }
    constructor() {
        super();
        _SpoOrgNewsSiteRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoOrgNewsSiteRemoveCommand_instances, "m", _SpoOrgNewsSiteRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoOrgNewsSiteRemoveCommand_instances, "m", _SpoOrgNewsSiteRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoOrgNewsSiteRemoveCommand_instances, "m", _SpoOrgNewsSiteRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeOrgNewsSite(logger, args.options.url);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove ${args.options.url} from the list of organizational news sites?` });
            if (result) {
                await this.removeOrgNewsSite(logger, args.options.url);
            }
        }
    }
    async removeOrgNewsSite(logger, url) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="64" ObjectPathId="63" /><Method Name="RemoveOrgNewsSite" Id="65" ObjectPathId="63"><Parameters><Parameter Type="String">${formatting.escapeXml(url)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="63" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
_SpoOrgNewsSiteRemoveCommand_instances = new WeakSet(), _SpoOrgNewsSiteRemoveCommand_initTelemetry = function _SpoOrgNewsSiteRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoOrgNewsSiteRemoveCommand_initOptions = function _SpoOrgNewsSiteRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '-f, --force'
    });
}, _SpoOrgNewsSiteRemoveCommand_initValidators = function _SpoOrgNewsSiteRemoveCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
};
export default new SpoOrgNewsSiteRemoveCommand();
//# sourceMappingURL=orgnewssite-remove.js.map