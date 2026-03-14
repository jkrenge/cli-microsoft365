import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoTenantRecycleBinItemListCommand extends SpoCommand {
    get name() {
        return commands.TENANT_RECYCLEBINITEM_LIST;
    }
    get description() {
        return 'Returns all modern and classic site collections in the tenant scoped recycle bin';
    }
    defaultProperties() {
        return ['DaysRemaining', 'DeletionTime', 'Url'];
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const res = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': res.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Url" ScalarProperty="true" /><Property Name="SiteId" ScalarProperty="true" /><Property Name="DaysRemaining" ScalarProperty="true" /><Property Name="Status" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetDeletedSitePropertiesFromSharePoint"><Parameters><Parameter Type="String">0</Parameter></Parameters></Method></ObjectPaths></Request>`
            };
            const processQuery = await request.post(requestOptions);
            const json = JSON.parse(processQuery);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            const results = json[json.length - 1];
            if (args.options.output !== 'json') {
                results._Child_Items_.forEach(s => {
                    s.DaysRemaining = Number(s.DaysRemaining);
                    s.DeletionTime = this.dateParser(s.DeletionTime);
                });
            }
            await logger.log(results._Child_Items_);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    dateParser(dateString) {
        const d = dateString.replace('/Date(', '').replace(')/', '').split(',').map(Number);
        return new Date(d[0], d[1], d[2], d[3], d[4], d[5], d[6]);
    }
}
export default new SpoTenantRecycleBinItemListCommand();
//# sourceMappingURL=tenant-recyclebinitem-list.js.map