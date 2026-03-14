var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpAiBuilderModelGetCommand_instances, _PpAiBuilderModelGetCommand_initTelemetry, _PpAiBuilderModelGetCommand_initOptions, _PpAiBuilderModelGetCommand_initOptionSets, _PpAiBuilderModelGetCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpAiBuilderModelGetCommand extends PowerPlatformCommand {
    get name() {
        return commands.AIBUILDERMODEL_GET;
    }
    get description() {
        return 'Get an AI builder model in the specified Power Platform environment.';
    }
    constructor() {
        super();
        _PpAiBuilderModelGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelGetCommand_instances, "m", _PpAiBuilderModelGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelGetCommand_instances, "m", _PpAiBuilderModelGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelGetCommand_instances, "m", _PpAiBuilderModelGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelGetCommand_instances, "m", _PpAiBuilderModelGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving an AI builder model '${args.options.id || args.options.name}'...`);
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const res = await this.getAiBuilderModel(dynamicsApiUrl, args.options);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAiBuilderModel(dynamicsApiUrl, options) {
        const requestOptions = {
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        if (options.id) {
            requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/msdyn_aimodels(${options.id})?$filter=iscustomizable/Value eq true`;
            const result = await request.get(requestOptions);
            return result;
        }
        requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/msdyn_aimodels?$filter=msdyn_name eq '${options.name}' and iscustomizable/Value eq true`;
        const result = await request.get(requestOptions);
        if (result.value.length === 0) {
            throw `The specified AI builder model '${options.name}' does not exist.`;
        }
        if (result.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('msdyn_aimodelid', result.value);
            return await cli.handleMultipleResultsFound(`Multiple AI builder models with name '${options.name}' found.`, resultAsKeyValuePair);
        }
        return result.value[0];
    }
}
_PpAiBuilderModelGetCommand_instances = new WeakSet(), _PpAiBuilderModelGetCommand_initTelemetry = function _PpAiBuilderModelGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpAiBuilderModelGetCommand_initOptions = function _PpAiBuilderModelGetCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--asAdmin'
    });
}, _PpAiBuilderModelGetCommand_initOptionSets = function _PpAiBuilderModelGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
}, _PpAiBuilderModelGetCommand_initValidators = function _PpAiBuilderModelGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpAiBuilderModelGetCommand();
//# sourceMappingURL=aibuildermodel-get.js.map