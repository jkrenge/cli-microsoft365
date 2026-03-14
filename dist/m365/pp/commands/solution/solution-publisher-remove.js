var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpSolutionPublisherRemoveCommand_instances, _PpSolutionPublisherRemoveCommand_initTelemetry, _PpSolutionPublisherRemoveCommand_initOptions, _PpSolutionPublisherRemoveCommand_initOptionSets, _PpSolutionPublisherRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import ppSolutionPublisherGetCommand from './solution-publisher-get.js';
class PpSolutionPublisherRemoveCommand extends PowerPlatformCommand {
    get name() {
        return commands.SOLUTION_PUBLISHER_REMOVE;
    }
    get description() {
        return 'Removes a specific publisher in the specified Power Platform environment.';
    }
    constructor() {
        super();
        _PpSolutionPublisherRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherRemoveCommand_instances, "m", _PpSolutionPublisherRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherRemoveCommand_instances, "m", _PpSolutionPublisherRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherRemoveCommand_instances, "m", _PpSolutionPublisherRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherRemoveCommand_instances, "m", _PpSolutionPublisherRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removes a publisher '${args.options.id || args.options.name}'...`);
        }
        if (args.options.force) {
            await this.deletePublisher(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove publisher '${args.options.id || args.options.name}'?` });
            if (result) {
                await this.deletePublisher(args);
            }
        }
    }
    async getPublisherId(args) {
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
        const output = await cli.executeCommandWithOutput(ppSolutionPublisherGetCommand, { options: { ...options, _: [] } });
        const getPublisherOutput = JSON.parse(output.stdout);
        return getPublisherOutput.publisherid;
    }
    async deletePublisher(args) {
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const publisherId = await this.getPublisherId(args);
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.1/publishers(${publisherId})`,
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
_PpSolutionPublisherRemoveCommand_instances = new WeakSet(), _PpSolutionPublisherRemoveCommand_initTelemetry = function _PpSolutionPublisherRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            asAdmin: !!args.options.asAdmin,
            force: !!args.options.force
        });
    });
}, _PpSolutionPublisherRemoveCommand_initOptions = function _PpSolutionPublisherRemoveCommand_initOptions() {
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
}, _PpSolutionPublisherRemoveCommand_initOptionSets = function _PpSolutionPublisherRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
}, _PpSolutionPublisherRemoveCommand_initValidators = function _PpSolutionPublisherRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpSolutionPublisherRemoveCommand();
//# sourceMappingURL=solution-publisher-remove.js.map