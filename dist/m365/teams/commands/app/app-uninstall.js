var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsAppUninstallCommand_instances, _TeamsAppUninstallCommand_initTelemetry, _TeamsAppUninstallCommand_initOptions, _TeamsAppUninstallCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsAppUninstallCommand extends GraphCommand {
    get name() {
        return commands.APP_UNINSTALL;
    }
    get description() {
        return 'Uninstalls an app from a Microsoft Team team';
    }
    constructor() {
        super();
        _TeamsAppUninstallCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsAppUninstallCommand_instances, "m", _TeamsAppUninstallCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsAppUninstallCommand_instances, "m", _TeamsAppUninstallCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsAppUninstallCommand_instances, "m", _TeamsAppUninstallCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const uninstallApp = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Uninstalling app with ID ${args.options.id} in team ${args.options.teamId}`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/teams/${args.options.teamId}/installedApps/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                }
            };
            try {
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await uninstallApp();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to uninstall the app with id ${args.options.id} from the Microsoft Teams team ${args.options.teamId}?` });
            if (result) {
                await uninstallApp();
            }
        }
    }
}
_TeamsAppUninstallCommand_instances = new WeakSet(), _TeamsAppUninstallCommand_initTelemetry = function _TeamsAppUninstallCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: args.options.force || false
        });
    });
}, _TeamsAppUninstallCommand_initOptions = function _TeamsAppUninstallCommand_initOptions() {
    this.options.unshift({
        option: '--id <id>'
    }, {
        option: '--teamId <teamId>'
    }, {
        option: '-f, --force'
    });
}, _TeamsAppUninstallCommand_initValidators = function _TeamsAppUninstallCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        return true;
    });
};
export default new TeamsAppUninstallCommand();
//# sourceMappingURL=app-uninstall.js.map