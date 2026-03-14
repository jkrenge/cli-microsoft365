import auth from '../../Auth.js';
import Command, { CommandError, globalOptionsZod } from '../../Command.js';
import commands from './commands.js';
export const options = globalOptionsZod.strict();
class LogoutCommand extends Command {
    get name() {
        return commands.LOGOUT;
    }
    get description() {
        return 'Log out from Microsoft 365';
    }
    get schema() {
        return options;
    }
    async commandAction(logger) {
        if (this.verbose) {
            await logger.logToStderr('Logging out from Microsoft 365...');
        }
        const deactivate = () => auth.connection.deactivate();
        try {
            await auth.clearConnectionInfo();
        }
        catch (error) {
            if (this.debug) {
                await logger.logToStderr(new CommandError(error));
            }
        }
        finally {
            deactivate();
        }
    }
    async action(logger, args) {
        try {
            await auth.restoreAuth();
        }
        catch (error) {
            throw new CommandError(error);
        }
        await this.initAction(args, logger);
        await this.commandAction(logger);
    }
}
export default new LogoutCommand();
//# sourceMappingURL=logout.js.map