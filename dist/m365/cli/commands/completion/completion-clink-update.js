import { autocomplete } from '../../../../autocomplete.js';
import AnonymousCommand from '../../../base/AnonymousCommand.js';
import commands from '../../commands.js';
class CliCompletionClinkUpdateCommand extends AnonymousCommand {
    get name() {
        return commands.COMPLETION_CLINK_UPDATE;
    }
    get description() {
        return 'Updates command completion for Clink (cmder)';
    }
    async commandAction(logger) {
        await logger.log(autocomplete.getClinkCompletion());
    }
}
export default new CliCompletionClinkUpdateCommand();
//# sourceMappingURL=completion-clink-update.js.map