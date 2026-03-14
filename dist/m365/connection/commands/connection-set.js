import auth from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandError, globalOptionsZod } from '../../../Command.js';
import z from 'zod';
export const options = z.strictObject({
    ...globalOptionsZod.shape,
    name: z.string().alias('n')
        .refine(async (name) => (await auth.getAllConnections()).some(c => c.name === name), {
        error: e => `Connection with name '${e.input}' does not exist.`
    }),
    newName: z.string()
        .refine(async (newName) => !(await auth.getAllConnections()).some(c => c.name === newName), {
        error: e => `Connection with name '${e.input}' already exists.`
    })
});
class ConnectionSetCommand extends Command {
    get name() {
        return commands.SET;
    }
    get description() {
        return 'Rename the specified connection';
    }
    get schema() {
        return options;
    }
    async commandAction(logger, args) {
        const connection = await auth.getConnection(args.options.name);
        if (this.verbose) {
            await logger.logToStderr(`Updating connection '${connection.identityName}', appId: ${connection.appId}, tenantId: ${connection.identityTenantId}...`);
        }
        await auth.updateConnection(connection, args.options.newName);
    }
    async action(logger, args) {
        try {
            await auth.restoreAuth();
        }
        catch (error) {
            throw new CommandError(error);
        }
        await this.initAction(args, logger);
        await this.commandAction(logger, args);
    }
}
export default new ConnectionSetCommand();
//# sourceMappingURL=connection-set.js.map