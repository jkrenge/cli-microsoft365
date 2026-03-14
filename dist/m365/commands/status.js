import { z } from 'zod';
import auth from '../../Auth.js';
import Command, { CommandError, globalOptionsZod } from '../../Command.js';
import commands from './commands.js';
export const options = z.strictObject({ ...globalOptionsZod.shape });
class StatusCommand extends Command {
    get name() {
        return commands.STATUS;
    }
    get description() {
        return 'Shows Microsoft 365 login status';
    }
    get schema() {
        return options;
    }
    async commandAction(logger) {
        if (auth.connection.active) {
            try {
                await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);
            }
            catch (err) {
                if (this.debug) {
                    await logger.logToStderr(err);
                }
                auth.connection.deactivate();
                throw new CommandError(`Your login has expired. Sign in again to continue. ${err.message}`);
            }
            const details = auth.getConnectionDetails(auth.connection);
            if (this.debug) {
                details.accessTokens = JSON.stringify(auth.connection.accessTokens, null, 2);
            }
            await logger.log(details);
        }
        else {
            const connections = await auth.getAllConnections();
            if (this.verbose) {
                const message = connections.length > 0
                    ? `Logged out, signed in connections available`
                    : 'Logged out';
                await logger.logToStderr(message);
            }
            else {
                const message = connections.length > 0
                    ? `Logged out, signed in connections available`
                    : 'Logged out';
                await logger.log(message);
            }
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
export default new StatusCommand();
//# sourceMappingURL=status.js.map