import fs from 'fs';
import { z } from 'zod';
import { cli } from '../../../cli/cli.js';
import { CommandError, globalOptionsZod } from '../../../Command.js';
import AnonymousCommand from '../../base/AnonymousCommand.js';
import commands from '../commands.js';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    force: z.boolean().optional().alias('f')
});
class ContextRemoveCommand extends AnonymousCommand {
    get name() {
        return commands.REMOVE;
    }
    get description() {
        return 'Removes the CLI for Microsoft 365 context in the current working folder';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            this.removeContext();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the context?` });
            if (result) {
                this.removeContext();
            }
        }
    }
    removeContext() {
        const filePath = '.m365rc.json';
        let m365rc = {};
        if (fs.existsSync(filePath)) {
            try {
                const fileContents = fs.readFileSync(filePath, 'utf8');
                if (fileContents) {
                    m365rc = JSON.parse(fileContents);
                }
            }
            catch (e) {
                throw new CommandError(`Error reading ${filePath}: ${e}. Please remove context info from ${filePath} manually.`);
            }
        }
        if (!m365rc.context) {
            return;
        }
        const keys = Object.keys(m365rc);
        if (keys.length === 1 && keys.indexOf('context') > -1) {
            try {
                fs.unlinkSync(filePath);
            }
            catch (e) {
                throw new CommandError(`Error removing ${filePath}: ${e}. Please remove ${filePath} manually.`);
            }
        }
        else {
            try {
                delete m365rc.context;
                fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
            }
            catch (e) {
                throw new CommandError(`Error writing ${filePath}: ${e}. Please remove context info from ${filePath} manually.`);
            }
        }
    }
}
export default new ContextRemoveCommand();
//# sourceMappingURL=context-remove.js.map