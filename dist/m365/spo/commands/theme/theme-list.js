import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoThemeListCommand extends SpoCommand {
    get name() {
        return commands.THEME_LIST;
    }
    get description() {
        return 'Retrieves the list of custom themes';
    }
    defaultProperties() {
        return ['name'];
    }
    async commandAction(logger) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving themes from tenant store...`);
            }
            const requestOptions = {
                url: `${spoAdminUrl}/_api/thememanager/GetTenantThemingOptions`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const rawRes = await request.post(requestOptions);
            await logger.log(rawRes.themePreviews || []);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoThemeListCommand();
//# sourceMappingURL=theme-list.js.map