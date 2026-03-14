var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ConnectionRemoveCommand_instances, _ConnectionRemoveCommand_initTelemetry, _ConnectionRemoveCommand_initOptions, _ConnectionRemoveCommand_initTypes;
import auth from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandError } from '../../../Command.js';
import { cli } from '../../../cli/cli.js';
class ConnectionRemoveCommand extends Command {
    get name() {
        return commands.REMOVE;
    }
    get description() {
        return 'Remove the specified connection';
    }
    constructor() {
        super();
        _ConnectionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _ConnectionRemoveCommand_instances, "m", _ConnectionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _ConnectionRemoveCommand_instances, "m", _ConnectionRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _ConnectionRemoveCommand_instances, "m", _ConnectionRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const deleteConnection = async () => {
            const connection = await auth.getConnection(args.options.name);
            if (this.verbose) {
                await logger.logToStderr(`Removing connection '${connection.identityName}', appId: ${connection.appId}, tenantId: ${connection.identityTenantId}...`);
            }
            await auth.removeConnectionInfo(connection, logger, this.debug);
        };
        if (args.options.force) {
            await deleteConnection();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the connection?` });
            if (result) {
                await deleteConnection();
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
        await this.commandAction(logger, args);
    }
}
_ConnectionRemoveCommand_instances = new WeakSet(), _ConnectionRemoveCommand_initTelemetry = function _ConnectionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _ConnectionRemoveCommand_initOptions = function _ConnectionRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-f, --force'
    });
}, _ConnectionRemoveCommand_initTypes = function _ConnectionRemoveCommand_initTypes() {
    this.types.string.push('name');
    this.types.boolean.push('force');
};
export default new ConnectionRemoveCommand();
//# sourceMappingURL=connection-remove.js.map