var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpCopilotGetCommand_instances, _PpCopilotGetCommand_initTelemetry, _PpCopilotGetCommand_initOptions, _PpCopilotGetCommand_initOptionSets, _PpCopilotGetCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class PpCopilotGetCommand extends PowerPlatformCommand {
    get name() {
        return commands.COPILOT_GET;
    }
    get description() {
        return 'Get information about the specified copilot';
    }
    constructor() {
        super();
        _PpCopilotGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpCopilotGetCommand_instances, "m", _PpCopilotGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpCopilotGetCommand_instances, "m", _PpCopilotGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpCopilotGetCommand_instances, "m", _PpCopilotGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpCopilotGetCommand_instances, "m", _PpCopilotGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving copilot '${args.options.id || args.options.name}'...`);
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const res = await this.getCopilot(dynamicsApiUrl, args.options);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getCopilot(dynamicsApiUrl, options) {
        const requestOptions = {
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        if (options.id) {
            requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/bots(${options.id})`;
            const result = await request.get(requestOptions);
            return result;
        }
        requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/bots?$filter=name eq '${formatting.encodeQueryParameter(options.name)}'`;
        const result = await request.get(requestOptions);
        if (result.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('botid', result.value);
            return await cli.handleMultipleResultsFound(`Multiple copilots with name '${options.name}' found.`, resultAsKeyValuePair);
        }
        if (result.value.length === 0) {
            throw `The specified copilot '${options.name}' does not exist.`;
        }
        return result.value[0];
    }
}
_PpCopilotGetCommand_instances = new WeakSet(), _PpCopilotGetCommand_initTelemetry = function _PpCopilotGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpCopilotGetCommand_initOptions = function _PpCopilotGetCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--asAdmin'
    });
}, _PpCopilotGetCommand_initOptionSets = function _PpCopilotGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
}, _PpCopilotGetCommand_initValidators = function _PpCopilotGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpCopilotGetCommand();
//# sourceMappingURL=copilot-get.js.map