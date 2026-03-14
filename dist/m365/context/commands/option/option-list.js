import fs from 'fs';
import { CommandError } from '../../../../Command.js';
import ContextCommand from '../../../base/ContextCommand.js';
import commands from '../../commands.js';
class ContextOptionListCommand extends ContextCommand {
    get name() {
        return commands.OPTION_LIST;
    }
    get description() {
        return 'List all options added to the context';
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving context options...`);
        }
        const filePath = '.m365rc.json';
        let m365rc = {};
        if (fs.existsSync(filePath)) {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Reading context file...`);
                }
                const fileContents = fs.readFileSync(filePath, 'utf8');
                if (fileContents) {
                    m365rc = JSON.parse(fileContents);
                }
            }
            catch (e) {
                throw new CommandError(`Error reading ${filePath}: ${e}. Please retrieve context options from ${filePath} manually.`);
            }
        }
        if (!m365rc.context) {
            throw new CommandError(`No context present`);
        }
        else {
            await logger.log(m365rc.context);
        }
    }
}
export default new ContextOptionListCommand();
//# sourceMappingURL=option-list.js.map