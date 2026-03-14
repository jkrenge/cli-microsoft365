var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ConnectionUseCommand_instances, _ConnectionUseCommand_initTelemetry, _ConnectionUseCommand_initOptions, _ConnectionUseCommand_initTypes;
import auth from '../../../Auth.js';
import commands from '../commands.js';
import Command, { CommandError } from '../../../Command.js';
import { formatting } from '../../../utils/formatting.js';
import { cli } from '../../../cli/cli.js';
class ConnectionUseCommand extends Command {
    get name() {
        return commands.USE;
    }
    get description() {
        return 'Activate the specified Microsoft 365 tenant connection';
    }
    constructor() {
        super();
        _ConnectionUseCommand_instances.add(this);
        __classPrivateFieldGet(this, _ConnectionUseCommand_instances, "m", _ConnectionUseCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _ConnectionUseCommand_instances, "m", _ConnectionUseCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _ConnectionUseCommand_instances, "m", _ConnectionUseCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        let connection;
        if (args.options.name) {
            connection = await auth.getConnection(args.options.name);
        }
        else {
            const connections = await auth.getAllConnections();
            connections.sort((a, b) => a.name.localeCompare(b.name));
            const keyValuePair = formatting.convertArrayToHashTable('name', connections);
            connection = await cli.handleMultipleResultsFound('Please select the connection you want to activate.', keyValuePair);
        }
        if (this.verbose) {
            await logger.logToStderr(`Switching to connection '${connection.identityName}', appId: ${connection.appId}, tenantId: ${connection.identityTenantId}...`);
        }
        await auth.switchToConnection(connection);
        const details = auth.getConnectionDetails(auth.connection);
        if (this.debug) {
            details.accessTokens = JSON.stringify(auth.connection.accessTokens, null, 2);
        }
        await logger.log(details);
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
_ConnectionUseCommand_instances = new WeakSet(), _ConnectionUseCommand_initTelemetry = function _ConnectionUseCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _ConnectionUseCommand_initOptions = function _ConnectionUseCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name [name]'
    });
}, _ConnectionUseCommand_initTypes = function _ConnectionUseCommand_initTypes() {
    this.types.string.push('name');
};
export default new ConnectionUseCommand();
//# sourceMappingURL=connection-use.js.map