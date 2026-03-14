import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteScriptListCommand extends SpoCommand {
    get name() {
        return commands.SITESCRIPT_LIST;
    }
    get description() {
        return 'Lists site script available for use with site designs';
    }
    async commandAction(logger) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const formDigest = await spo.getRequestDigest(spoUrl);
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts'`,
                headers: {
                    'X-RequestDigest': formDigest.FormDigestValue,
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            await logger.log(res.value);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoSiteScriptListCommand();
//# sourceMappingURL=sitescript-list.js.map