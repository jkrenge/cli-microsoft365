import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
const options = globalOptionsZod.strict();
class SpoBrandCenterSettingsListCommand extends SpoCommand {
    get name() {
        return commands.BRANDCENTER_SETTINGS_LIST;
    }
    get description() {
        return 'Lists the brand center configuration';
    }
    get schema() {
        return options;
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving brand center configuration...`);
        }
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.verbose);
            const requestOptions = {
                url: `${spoUrl}/_api/Brandcenter/Configuration`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new SpoBrandCenterSettingsListCommand();
//# sourceMappingURL=brandcenter-settings-list.js.map