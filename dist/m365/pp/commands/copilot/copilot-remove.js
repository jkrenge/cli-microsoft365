var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpCopilotRemoveCommand_instances, _PpCopilotRemoveCommand_initTelemetry, _PpCopilotRemoveCommand_initOptions, _PpCopilotRemoveCommand_initOptionSets, _PpCopilotRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import ppCopilotGetCommand from './copilot-get.js';
class PpCopilotRemoveCommand extends PowerPlatformCommand {
    get name() {
        return commands.COPILOT_REMOVE;
    }
    get description() {
        return 'Removes the specified copilot';
    }
    constructor() {
        super();
        _PpCopilotRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpCopilotRemoveCommand_instances, "m", _PpCopilotRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpCopilotRemoveCommand_instances, "m", _PpCopilotRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpCopilotRemoveCommand_instances, "m", _PpCopilotRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpCopilotRemoveCommand_instances, "m", _PpCopilotRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing copilot '${args.options.id || args.options.name}'...`);
        }
        if (args.options.force) {
            await this.deleteCopilot(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove copilot '${args.options.id || args.options.name}'?` });
            if (result) {
                await this.deleteCopilot(args);
            }
        }
    }
    async getCopilotId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const options = {
            environmentName: args.options.environmentName,
            name: args.options.name,
            output: 'json',
            debug: this.debug,
            verbose: this.verbose
        };
        const output = await cli.executeCommandWithOutput(ppCopilotGetCommand, { options: { ...options, _: [] } });
        const getBotOutput = JSON.parse(output.stdout);
        return getBotOutput.botid;
    }
    async deleteCopilot(args) {
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const botId = await this.getCopilotId(args);
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.1/bots(${botId})/Microsoft.Dynamics.CRM.PvaDeleteBot?tag=deprovisionbotondelete`,
                headers: {
                    accept: 'application/json',
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpCopilotRemoveCommand_instances = new WeakSet(), _PpCopilotRemoveCommand_initTelemetry = function _PpCopilotRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            asAdmin: !!args.options.asAdmin,
            force: !!args.options.force
        });
    });
}, _PpCopilotRemoveCommand_initOptions = function _PpCopilotRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--asAdmin'
    }, {
        option: '-f, --force'
    });
}, _PpCopilotRemoveCommand_initOptionSets = function _PpCopilotRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
}, _PpCopilotRemoveCommand_initValidators = function _PpCopilotRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpCopilotRemoveCommand();
//# sourceMappingURL=copilot-remove.js.map