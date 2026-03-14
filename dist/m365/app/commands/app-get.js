import { entraApp } from '../../../utils/entraApp.js';
import AppCommand, { appCommandOptions } from '../../base/AppCommand.js';
import commands from '../commands.js';
class AppGetCommand extends AppCommand {
    get name() {
        return commands.GET;
    }
    get description() {
        return 'Retrieves information about the current Microsoft Entra app';
    }
    get schema() {
        return appCommandOptions;
    }
    async commandAction(logger, args) {
        try {
            const app = await entraApp.getAppRegistrationByAppId(args.options.appId);
            await logger.log(app);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
export default new AppGetCommand();
//# sourceMappingURL=app-get.js.map