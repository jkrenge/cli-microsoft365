import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHideDefaultThemesGetCommand extends SpoCommand {
    get name() {
        return commands.HIDEDEFAULTTHEMES_GET;
    }
    get description() {
        return 'Gets the current value of the HideDefaultThemes setting';
    }
    async commandAction(logger) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Getting the current value of the HideDefaultThemes setting...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_api/thememanager/GetHideDefaultThemes`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
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
export default new SpoHideDefaultThemesGetCommand();
//# sourceMappingURL=hidedefaultthemes-get.js.map