import { app } from '../../utils/app.js';
import AnonymousCommand from '../base/AnonymousCommand.js';
import commands from './commands.js';
class VersionCommand extends AnonymousCommand {
    get name() {
        return commands.VERSION;
    }
    get description() {
        return 'Shows CLI for Microsoft 365 version';
    }
    async commandAction(logger) {
        await logger.log(`v${app.packageJson().version}`);
    }
}
export default new VersionCommand();
//# sourceMappingURL=version.js.map