import assert from 'assert';
import { z } from 'zod';
import auth from '../../../Auth.js';
import Command, { CommandError, globalOptionsZod } from '../../../Command.js';
import commands from '../commands.js';
export const options = z.strictObject({ ...globalOptionsZod.shape });
class ConnectionListCommand extends Command {
    get name() {
        return commands.LIST;
    }
    get description() {
        return 'Show the list of available connections';
    }
    get schema() {
        return options;
    }
    defaultProperties() {
        return ['name', 'connectedAs', 'authType', 'active'];
    }
    async commandAction(logger) {
        const availableConnections = await auth.getAllConnections();
        const output = availableConnections.map(connection => {
            const isCurrentConnection = connection.name === auth.connection?.name;
            return {
                name: connection.name,
                connectedAs: connection.identityName,
                authType: connection.authType,
                active: isCurrentConnection
            };
        }).sort((a, b) => {
            // Asserting name because it is optional, but required at this point.
            assert(a.name !== undefined);
            assert(b.name !== undefined);
            const aName = a.name;
            const bName = b.name;
            return aName.localeCompare(bName);
        });
        await logger.log(output);
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
export default new ConnectionListCommand();
//# sourceMappingURL=connection-list.js.map