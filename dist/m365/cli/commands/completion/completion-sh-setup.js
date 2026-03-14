import { autocomplete } from '../../../../autocomplete.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
class CliCompletionShSetupCommand extends AnonymousCommand {
    get name() {
        return commands.COMPLETION_SH_SETUP;
    }
    get description() {
        return 'Sets up command completion for Zsh, Bash and Fish';
    }
    async commandAction(logger) {
        if (this.debug) {
            await logger.logToStderr('Generating command completion...');
        }
        autocomplete.generateShCompletion();
        if (this.debug) {
            await logger.logToStderr('Registering command completion with the shell...');
        }
        autocomplete.setupShCompletion();
        await logger.log('Command completion successfully registered. Restart your shell to load the completion');
    }
}
export default new CliCompletionShSetupCommand();
//# sourceMappingURL=completion-sh-setup.js.map