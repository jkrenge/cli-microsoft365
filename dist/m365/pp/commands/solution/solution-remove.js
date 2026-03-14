var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpSolutionRemoveCommand_instances, _PpSolutionRemoveCommand_initTelemetry, _PpSolutionRemoveCommand_initOptions, _PpSolutionRemoveCommand_initOptionSets, _PpSolutionRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpSolutionRemoveCommand extends PowerPlatformCommand {
    get name() {
        return commands.SOLUTION_REMOVE;
    }
    get description() {
        return 'Removes a specific solution in the specified Power Platform environment.';
    }
    constructor() {
        super();
        _PpSolutionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpSolutionRemoveCommand_instances, "m", _PpSolutionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpSolutionRemoveCommand_instances, "m", _PpSolutionRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpSolutionRemoveCommand_instances, "m", _PpSolutionRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpSolutionRemoveCommand_instances, "m", _PpSolutionRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing solution '${args.options.id || args.options.name}'...`);
        }
        if (args.options.force) {
            await this.deleteSolution(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove solution '${args.options.id || args.options.name}'?` });
            if (result) {
                await this.deleteSolution(args);
            }
        }
    }
    async getSolutionId(args, dynamicsApiUrl) {
        if (args.options.id) {
            return args.options.id;
        }
        const solution = await powerPlatform.getSolutionByName(dynamicsApiUrl, args.options.name);
        return solution.solutionid;
    }
    async deleteSolution(args) {
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const solutionId = await this.getSolutionId(args, dynamicsApiUrl);
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.1/solutions(${solutionId})`,
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
_PpSolutionRemoveCommand_instances = new WeakSet(), _PpSolutionRemoveCommand_initTelemetry = function _PpSolutionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            asAdmin: !!args.options.asAdmin,
            force: !!args.options.force
        });
    });
}, _PpSolutionRemoveCommand_initOptions = function _PpSolutionRemoveCommand_initOptions() {
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
}, _PpSolutionRemoveCommand_initOptionSets = function _PpSolutionRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
}, _PpSolutionRemoveCommand_initValidators = function _PpSolutionRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpSolutionRemoveCommand();
//# sourceMappingURL=solution-remove.js.map