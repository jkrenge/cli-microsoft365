var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpAiBuilderModelRemoveCommand_instances, _PpAiBuilderModelRemoveCommand_initTelemetry, _PpAiBuilderModelRemoveCommand_initOptions, _PpAiBuilderModelRemoveCommand_initOptionSets, _PpAiBuilderModelRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import ppAiBuilderModelGetCommand from './aibuildermodel-get.js';
class PpAiBuilderModelRemoveCommand extends PowerPlatformCommand {
    get name() {
        return commands.AIBUILDERMODEL_REMOVE;
    }
    get description() {
        return 'Removes an AI builder model in the specified Power Platform environment.';
    }
    constructor() {
        super();
        _PpAiBuilderModelRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelRemoveCommand_instances, "m", _PpAiBuilderModelRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelRemoveCommand_instances, "m", _PpAiBuilderModelRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelRemoveCommand_instances, "m", _PpAiBuilderModelRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelRemoveCommand_instances, "m", _PpAiBuilderModelRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing AI builder model '${args.options.id || args.options.name}'...`);
        }
        if (args.options.force) {
            await this.deleteAiBuilderModel(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove AI builder model '${args.options.id || args.options.name}'?` });
            if (result) {
                await this.deleteAiBuilderModel(args);
            }
        }
    }
    async getAiBuilderModelId(args) {
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
        const output = await cli.executeCommandWithOutput(ppAiBuilderModelGetCommand, { options: { ...options, _: [] } });
        const getAiBuilderModelOutput = JSON.parse(output.stdout);
        return getAiBuilderModelOutput.msdyn_aimodelid;
    }
    async deleteAiBuilderModel(args) {
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const aiBuilderModelId = await this.getAiBuilderModelId(args);
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.1/msdyn_aimodels(${aiBuilderModelId})`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpAiBuilderModelRemoveCommand_instances = new WeakSet(), _PpAiBuilderModelRemoveCommand_initTelemetry = function _PpAiBuilderModelRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            asAdmin: !!args.options.asAdmin,
            force: !!args.options.force
        });
    });
}, _PpAiBuilderModelRemoveCommand_initOptions = function _PpAiBuilderModelRemoveCommand_initOptions() {
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
}, _PpAiBuilderModelRemoveCommand_initOptionSets = function _PpAiBuilderModelRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
}, _PpAiBuilderModelRemoveCommand_initValidators = function _PpAiBuilderModelRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpAiBuilderModelRemoveCommand();
//# sourceMappingURL=aibuildermodel-remove.js.map