import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteDesignListCommand extends SpoCommand {
    get name() {
        return commands.SITEDESIGN_LIST;
    }
    get description() {
        return 'Lists available site designs for creating modern sites';
    }
    defaultProperties() {
        return ['Id', 'IsDefault', 'Title', 'Version', 'WebTemplate'];
    }
    async commandAction(logger) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const requestOptions = {
                url: `${spoUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
                headers: {
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
export default new SpoSiteDesignListCommand();
//# sourceMappingURL=sitedesign-list.js.map