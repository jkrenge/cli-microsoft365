import ContextCommand from '../../base/ContextCommand.js';
import commands from '../commands.js';
class ContextInitCommand extends ContextCommand {
    get name() {
        return commands.INIT;
    }
    get description() {
        return 'Initiates CLI for Microsoft 365 context in the current working folder';
    }
    async commandAction() {
        await this.saveContextInfo({});
    }
}
export default new ContextInitCommand();
//# sourceMappingURL=context-init.js.map