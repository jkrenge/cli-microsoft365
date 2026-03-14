var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaAppRemoveCommand_instances, _PaAppRemoveCommand_initTelemetry, _PaAppRemoveCommand_initOptions, _PaAppRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
class PaAppRemoveCommand extends PowerAppsCommand {
    get name() {
        return commands.APP_REMOVE;
    }
    get description() {
        return 'Removes the specified Microsoft Power App';
    }
    constructor() {
        super();
        _PaAppRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PaAppRemoveCommand_instances, "m", _PaAppRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PaAppRemoveCommand_instances, "m", _PaAppRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaAppRemoveCommand_instances, "m", _PaAppRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing Microsoft Power App ${args.options.name}...`);
        }
        const removePaApp = async () => {
            let endpoint = `${this.resource}/providers/Microsoft.PowerApps`;
            if (args.options.asAdmin) {
                endpoint += `/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName)}`;
            }
            endpoint += `/apps/${formatting.encodeQueryParameter(args.options.name)}?api-version=2017-08-01`;
            const requestOptions = {
                url: endpoint,
                fullResponse: true,
                headers: {
                    accept: 'application/json'
                },
                responseType: 'json'
            };
            try {
                await request.delete(requestOptions);
            }
            catch (err) {
                if (err.response && err.response.status === 403) {
                    throw new CommandError(`App '${args.options.name}' does not exist`);
                }
                else {
                    this.handleRejectedODataJsonPromise(err);
                }
            }
        };
        if (args.options.force) {
            await removePaApp();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Microsoft Power App ${args.options.name}?` });
            if (result) {
                await removePaApp();
            }
        }
    }
}
_PaAppRemoveCommand_instances = new WeakSet(), _PaAppRemoveCommand_initTelemetry = function _PaAppRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: typeof args.options.force !== 'undefined',
            asAdmin: !!args.options.asAdmin,
            environmentName: typeof args.options.environmentName !== 'undefined'
        });
    });
}, _PaAppRemoveCommand_initOptions = function _PaAppRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-f, --force'
    }, {
        option: '--asAdmin'
    }, {
        option: '-e, --environmentName [environmentName]'
    });
}, _PaAppRemoveCommand_initValidators = function _PaAppRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.name)) {
            return `${args.options.name} is not a valid GUID`;
        }
        if (args.options.asAdmin && !args.options.environmentName) {
            return 'When specifying the asAdmin option, the environment option is required as well.';
        }
        if (args.options.environmentName && !args.options.asAdmin) {
            return 'When specifying the environment option, the asAdmin option is required as well.';
        }
        return true;
    });
};
export default new PaAppRemoveCommand();
//# sourceMappingURL=app-remove.js.map