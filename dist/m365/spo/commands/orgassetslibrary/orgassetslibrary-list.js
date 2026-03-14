import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoOrgAssetsLibraryListCommand extends SpoCommand {
    get name() {
        return commands.ORGASSETSLIBRARY_LIST;
    }
    get description() {
        return 'List all libraries that are assigned as asset library';
    }
    async commandAction(logger) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const reqDigest = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': reqDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Method Name="GetOrgAssets" Id="6" ObjectPathId="3" /></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                const orgAssetsResponse = json[json.length - 1];
                if (orgAssetsResponse === null || orgAssetsResponse.OrgAssetsLibraries === undefined) {
                    await logger.log("No libraries in Organization Assets");
                }
                else {
                    const orgAssets = {
                        Url: orgAssetsResponse.Url.DecodedUrl,
                        Libraries: orgAssetsResponse.OrgAssetsLibraries._Child_Items_.map(t => {
                            return {
                                DisplayName: t.DisplayName,
                                LibraryUrl: t.LibraryUrl.DecodedUrl,
                                ListId: t.ListId,
                                ThumbnailUrl: t.ThumbnailUrl !== null ? t.ThumbnailUrl.DecodedUrl : null
                            };
                        })
                    };
                    await logger.log(orgAssets);
                }
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
export default new SpoOrgAssetsLibraryListCommand();
//# sourceMappingURL=orgassetslibrary-list.js.map