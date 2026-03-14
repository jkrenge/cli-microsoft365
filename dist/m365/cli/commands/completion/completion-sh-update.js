import { autocomplete } from '../../../../autocomplete.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
class CliCompletionShUpdateCommand extends AnonymousCommand {
    get name() {
        return commands.COMPLETION_SH_UPDATE;
    }
    get description() {
        return 'Updates command completion for Zsh, Bash and Fish';
    }
    async commandAction(logger) {
        if (this.debug) {
            await logger.logToStderr('Generating command completion...');
        }
        autocomplete.generateShCompletion();
    }
}
export default new CliCompletionShUpdateCommand();
//# sourceMappingURL=completion-sh-update.js.map