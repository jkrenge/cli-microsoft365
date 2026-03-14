import fs from 'fs';
import { CommandError } from '../../Command.js';
import AnonymousCommand from './AnonymousCommand.js';
export default class ContextCommand extends AnonymousCommand {
    saveContextInfo(context) {
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
                throw new CommandError(`Error reading ${filePath}: ${e}. Please add context info to ${filePath} manually.`);
            }
        }
        if (!m365rc.context) {
            m365rc.context = context;
            try {
                fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
            }
            catch (e) {
                throw new CommandError(`Error writing ${filePath}: ${e}. Please add context info to ${filePath} manually.`);
            }
        }
    }
}
//# sourceMappingURL=ContextCommand.js.map