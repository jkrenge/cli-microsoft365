import { autocomplete } from '../../../../autocomplete.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
class CliCompletionPwshUpdateCommand extends AnonymousCommand {
    get name() {
        return commands.COMPLETION_PWSH_UPDATE;
    }
    get description() {
        return 'Updates command completion for PowerShell';
    }
    async commandAction(logger) {
        if (this.debug) {
            await logger.logToStderr('Generating command completion...');
        }
        autocomplete.generateShCompletion();
    }
}
export default new CliCompletionPwshUpdateCommand();
//# sourceMappingURL=completion-pwsh-update.js.map