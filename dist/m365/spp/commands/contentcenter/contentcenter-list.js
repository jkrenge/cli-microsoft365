import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SppContentCenterListCommand extends SpoCommand {
    get name() {
        return commands.CONTENTCENTER_LIST;
    }
    get description() {
        return 'Gets information about the SharePoint Premium content centers';
    }
    defaultProperties() {
        return ['Title', 'Url'];
    }
    async commandAction(logger) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving list of content centers...`);
            }
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const allContentCenters = await this.getContentCenters(spoAdminUrl, logger);
            await logger.log(allContentCenters);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getContentCenters(spoAdminUrl, logger) {
        const allSites = [];
        let currentStartIndex = '0';
        const res = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, this.debug);
        do {
            const requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String"></Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">0</Property><Property Name="StartIndex" Type="String">${currentStartIndex}</Property><Property Name="Template" Type="String">CONTENTCTR#0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`;
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: requestBody
            };
            const response = await request.post(requestOptions);
            const json = JSON.parse(response);
            const responseContent = json[0];
            if (responseContent.ErrorInfo) {
                throw responseContent.ErrorInfo.ErrorMessage;
            }
            const sites = json[json.length - 1];
            allSites.push(...sites._Child_Items_);
            currentStartIndex = sites.NextStartIndexFromSharePoint;
        } while (currentStartIndex);
        return allSites;
    }
}
export default new SppContentCenterListCommand();
//# sourceMappingURL=contentcenter-list.js.map