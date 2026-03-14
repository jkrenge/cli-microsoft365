import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantSettingsListCommand extends SpoCommand {
    get name() {
        return commands.TENANT_SETTINGS_LIST;
    }
    get description() {
        return 'Lists the global tenant settings';
    }
    async commandAction(logger) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties><Property Name="HideDefaultThemes" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const result = json[4];
            delete result['_ObjectIdentity_'];
            delete result['_ObjectType_'];
            // map integers to their enums
            const sharingLinkType = ['None', 'Direct', 'Internal', 'AnonymousAccess'];
            const sharingCapabilities = ['Disabled', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly'];
            const sharingDomainRestrictionModes = ['None', 'AllowList', 'BlockList'];
            const sharingState = ['Unspecified', 'On', 'Off'];
            const anonymousLinkType = ['None', 'View', 'Edit'];
            const sharingPermissionType = ['None', 'View', 'Edit'];
            const sPOConditionalAccessPolicyType = ['AllowFullAccess', 'AllowLimitedAccess', 'BlockAccess'];
            const specialCharactersState = ['NoPreference', 'Allowed', 'Disallowed'];
            const sPOLimitedAccessFileType = ['OfficeOnlineFilesOnly', 'WebPreviewableFiles', 'OtherFiles'];
            result['SharingCapability'] = sharingCapabilities[result['SharingCapability']];
            result['SharingDomainRestrictionMode'] = sharingDomainRestrictionModes[result['SharingDomainRestrictionMode']];
            result['ODBMembersCanShare'] = sharingState[result['ODBMembersCanShare']];
            result['ODBAccessRequests'] = sharingState[result['ODBAccessRequests']];
            result['DefaultSharingLinkType'] = sharingLinkType[result['DefaultSharingLinkType']];
            result['FileAnonymousLinkType'] = anonymousLinkType[result['FileAnonymousLinkType']];
            result['FolderAnonymousLinkType'] = anonymousLinkType[result['FolderAnonymousLinkType']];
            result['DefaultLinkPermission'] = sharingPermissionType[result['DefaultLinkPermission']];
            result['ConditionalAccessPolicy'] = sPOConditionalAccessPolicyType[result['ConditionalAccessPolicy']];
            result['SpecialCharactersStateInFileFolderNames'] = specialCharactersState[result['SpecialCharactersStateInFileFolderNames']];
            result['LimitedAccessFileType'] = sPOLimitedAccessFileType[result['LimitedAccessFileType']];
            await logger.log(result);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoTenantSettingsListCommand();
//# sourceMappingURL=tenant-settings-list.js.map