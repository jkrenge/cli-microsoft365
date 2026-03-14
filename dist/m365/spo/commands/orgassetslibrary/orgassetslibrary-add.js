var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoOrgAssetsLibraryAddCommand_instances, _a, _SpoOrgAssetsLibraryAddCommand_initTelemetry, _SpoOrgAssetsLibraryAddCommand_initOptions, _SpoOrgAssetsLibraryAddCommand_initValidators;
import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
var OrgAssetType;
(function (OrgAssetType) {
    OrgAssetType[OrgAssetType["ImageDocumentLibrary"] = 1] = "ImageDocumentLibrary";
    OrgAssetType[OrgAssetType["OfficeTemplateLibrary"] = 2] = "OfficeTemplateLibrary";
    OrgAssetType[OrgAssetType["OfficeFontLibrary"] = 4] = "OfficeFontLibrary";
})(OrgAssetType || (OrgAssetType = {}));
class SpoOrgAssetsLibraryAddCommand extends SpoCommand {
    get name() {
        return commands.ORGASSETSLIBRARY_ADD;
    }
    get description() {
        return 'Promotes an existing library to become an organization assets library';
    }
    constructor() {
        super();
        _SpoOrgAssetsLibraryAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoOrgAssetsLibraryAddCommand_instances, "m", _SpoOrgAssetsLibraryAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoOrgAssetsLibraryAddCommand_instances, "m", _SpoOrgAssetsLibraryAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoOrgAssetsLibraryAddCommand_instances, "m", _SpoOrgAssetsLibraryAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const cdnTypeString = args.options.cdnType || 'Private';
        const cdnType = cdnTypeString === 'Private' ? 1 : 0;
        const thumbnailSchema = typeof args.options.thumbnailUrl === 'undefined' ? `<Parameter Type="Null" />` : `<Parameter Type="String">${args.options.thumbnailUrl}</Parameter>`;
        try {
            const orgAssetType = this.getOrgAssetType(args.options.orgAssetType);
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="AddToOrgAssetsLibAndCdnWithType" Id="11" ObjectPathId="8"><Parameters><Parameter Type="Enum">${cdnType}</Parameter><Parameter Type="String">${args.options.libraryUrl}</Parameter>${thumbnailSchema}<Parameter Type="Enum">${orgAssetType}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="8" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
    getOrgAssetType(orgAssetType) {
        switch (orgAssetType) {
            case 'OfficeTemplateLibrary':
                return OrgAssetType.OfficeTemplateLibrary;
            case 'OfficeFontLibrary':
                return OrgAssetType.OfficeFontLibrary;
            default:
                return OrgAssetType.ImageDocumentLibrary;
        }
    }
}
_a = SpoOrgAssetsLibraryAddCommand, _SpoOrgAssetsLibraryAddCommand_instances = new WeakSet(), _SpoOrgAssetsLibraryAddCommand_initTelemetry = function _SpoOrgAssetsLibraryAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            cdnType: args.options.cdnType || 'Private',
            thumbnailUrl: typeof args.options.thumbnailUrl !== 'undefined',
            orgAssetType: args.options.orgAssetType
        });
    });
}, _SpoOrgAssetsLibraryAddCommand_initOptions = function _SpoOrgAssetsLibraryAddCommand_initOptions() {
    this.options.unshift({
        option: '--libraryUrl <libraryUrl>'
    }, {
        option: '--thumbnailUrl [thumbnailUrl]'
    }, {
        option: '--cdnType [cdnType]',
        autocomplete: _a.cdnTypes
    }, {
        option: '--orgAssetType [orgAssetType]',
        autocomplete: _a.orgAssetTypes
    });
}, _SpoOrgAssetsLibraryAddCommand_initValidators = function _SpoOrgAssetsLibraryAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidThumbnailUrl = validation.isValidSharePointUrl(args.options.thumbnailUrl);
        if (typeof args.options.thumbnailUrl !== 'undefined' && isValidThumbnailUrl !== true) {
            return isValidThumbnailUrl;
        }
        if (args.options.cdnType && _a.cdnTypes.indexOf(args.options.cdnType) < 0) {
            return `${args.options.cdnType} is not a valid value for cdnType. Valid values are ${_a.cdnTypes.join(', ')}`;
        }
        if (args.options.orgAssetType && _a.orgAssetTypes.indexOf(args.options.orgAssetType) < 0) {
            return `${args.options.orgAssetType} is not a valid value for orgAssetType. Valid values are ${_a.orgAssetTypes.join(', ')}`;
        }
        return validation.isValidSharePointUrl(args.options.libraryUrl);
    });
};
SpoOrgAssetsLibraryAddCommand.orgAssetTypes = ['ImageDocumentLibrary', 'OfficeTemplateLibrary', 'OfficeFontLibrary'];
SpoOrgAssetsLibraryAddCommand.cdnTypes = ['Public', 'Private'];
export default new SpoOrgAssetsLibraryAddCommand();
//# sourceMappingURL=orgassetslibrary-add.js.map