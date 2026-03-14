var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoOrgAssetsLibraryRemoveCommand_instances, _SpoOrgAssetsLibraryRemoveCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoOrgAssetsLibraryRemoveCommand extends SpoCommand {
    get name() {
        return commands.ORGASSETSLIBRARY_REMOVE;
    }
    get description() {
        return 'Removes a library that was designated as a central location for organization assets across the tenant.';
    }
    constructor() {
        super();
        _SpoOrgAssetsLibraryRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoOrgAssetsLibraryRemoveCommand_instances, "m", _SpoOrgAssetsLibraryRemoveCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeLibrary(logger, args.options.libraryUrl);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the library ${args.options.libraryUrl} as a central location for organization assets?` });
            if (result) {
                await this.removeLibrary(logger, args.options.libraryUrl);
            }
        }
    }
    async removeLibrary(logger, libraryUrl) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="RemoveFromOrgAssets" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">${libraryUrl}</Parameter><Parameter Type="Guid">{00000000-0000-0000-0000-000000000000}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                await logger.log(json[json.length - 1]);
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
_SpoOrgAssetsLibraryRemoveCommand_instances = new WeakSet(), _SpoOrgAssetsLibraryRemoveCommand_initOptions = function _SpoOrgAssetsLibraryRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--libraryUrl <libraryUrl>'
    }, {
        option: '-f, --force'
    });
};
export default new SpoOrgAssetsLibraryRemoveCommand();
//# sourceMappingURL=orgassetslibrary-remove.js.map