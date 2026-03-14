import config from '../../../../config.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoOrgNewsSiteListCommand extends SpoCommand {
    get name() {
        return commands.ORGNEWSSITE_LIST;
    }
    get description() {
        return 'Lists all organizational news sites';
    }
    async commandAction(logger) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const resDigest = await spo.getRequestDigest(spoAdminUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
                headers: {
                    'X-RequestDigest': resDigest.FormDigestValue
                },
                data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="58" ObjectPathId="57" /><Method Name="GetOrgNewsSites" Id="59" ObjectPathId="57" /></Actions><ObjectPaths><Constructor Id="57" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
            };
            const res = await request.post(requestOptions);
            const json = JSON.parse(res);
            const response = json[0];
            if (response.ErrorInfo) {
                throw response.ErrorInfo.ErrorMessage;
            }
            else {
                const results = json[json.length - 1];
                await logger.log(results);
            }
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
}
export default new SpoOrgNewsSiteListCommand();
//# sourceMappingURL=orgnewssite-list.js.map