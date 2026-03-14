import { cli } from '../../cli/cli.js';
import { globalOptionsZod } from '../../Command.js';
import { settingsNames } from '../../settingsNames.js';
import { app } from '../../utils/app.js';
import { browserUtil } from '../../utils/browserUtil.js';
import AnonymousCommand from '../base/AnonymousCommand.js';
import commands from './commands.js';
export const options = globalOptionsZod.strict();
class DocsCommand extends AnonymousCommand {
    get name() {
        return commands.DOCS;
    }
    get description() {
        return 'Returns the CLI for Microsoft 365 docs webpage URL';
    }
    get schema() {
        return options;
    }
    async commandAction(logger) {
        await logger.log(app.packageJson().homepage);
        if (cli.getSettingWithDefaultValue(settingsNames.autoOpenLinksInBrowser, false) === false) {
            return;
        }
        await browserUtil.open(app.packageJson().homepage);
    }
}
export default new DocsCommand();
//# sourceMappingURL=docs.js.map