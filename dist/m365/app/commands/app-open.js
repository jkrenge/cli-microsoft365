import { z } from 'zod';
import { cli } from '../../../cli/cli.js';
import { settingsNames } from '../../../settingsNames.js';
import { browserUtil } from '../../../utils/browserUtil.js';
import AppCommand, { appCommandOptions } from '../../base/AppCommand.js';
import commands from '../commands.js';
export const options = z.strictObject({
    ...appCommandOptions.shape,
    preview: z.boolean().optional().default(false)
});
class AppOpenCommand extends AppCommand {
    get name() {
        return commands.OPEN;
    }
    get description() {
        return 'Opens Microsoft Entra app in the Microsoft Entra ID portal';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        try {
            await this.logOrOpenUrl(args, logger);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async logOrOpenUrl(args, logger) {
        const previewPrefix = args.options.preview === true ? "preview." : "";
        const url = `https://${previewPrefix}portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${this.appId}/isMSAApp/`;
        if (cli.getSettingWithDefaultValue(settingsNames.autoOpenLinksInBrowser, false) === false) {
            await logger.log(`Use a web browser to open the page ${url}`);
            return;
        }
        await logger.log(`Opening the following page in your browser: ${url}`);
        await browserUtil.open(url);
    }
}
export default new AppOpenCommand();
//# sourceMappingURL=app-open.js.map