import fs from 'fs';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { CommandError, globalOptionsZod } from '../../../../Command.js';
import ContextCommand from '../../../base/ContextCommand.js';
import commands from '../../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    name: z.string().alias('n'),
    force: z.boolean().optional().alias('f')
});
class ContextOptionRemoveCommand extends ContextCommand {
    get name() {
        return commands.OPTION_REMOVE;
    }
    get description() {
        return 'Removes an already available name from local context file.';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing context option '${args.options.name}'...`);
        }
        if (args.options.force) {
            await this.removeContextOption(args.options.name, logger);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the context option ${args.options.name}?` });
            if (result) {
                await this.removeContextOption(args.options.name, logger);
            }
        }
    }
    async removeContextOption(name, logger) {
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
                throw new CommandError(`Error reading ${filePath}: ${e}. Please remove context option ${name} from ${filePath} manually.`);
            }
        }
        if (!m365rc.context || !m365rc.context[name]) {
            throw new CommandError(`There is no option ${name} in the context info`);
        }
        else {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Removing context option ${name} from the context file...`);
                }
                delete m365rc.context[name];
                fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
            }
            catch (e) {
                throw new CommandError(`Error writing ${filePath}: ${e}. Please remove context option ${name} from ${filePath} manually.`);
            }
        }
    }
}
export default new ContextOptionRemoveCommand();
//# sourceMappingURL=option-remove.js.map