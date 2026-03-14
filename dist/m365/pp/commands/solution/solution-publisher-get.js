var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpSolutionPublisherGetCommand_instances, _PpSolutionPublisherGetCommand_initTelemetry, _PpSolutionPublisherGetCommand_initOptions, _PpSolutionPublisherGetCommand_initOptionSets, _PpSolutionPublisherGetCommand_initValidators;
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpSolutionPublisherGetCommand extends PowerPlatformCommand {
    get name() {
        return commands.SOLUTION_PUBLISHER_GET;
    }
    get description() {
        return 'Get information about the specified publisher in a given environment.';
    }
    constructor() {
        super();
        _PpSolutionPublisherGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherGetCommand_instances, "m", _PpSolutionPublisherGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherGetCommand_instances, "m", _PpSolutionPublisherGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherGetCommand_instances, "m", _PpSolutionPublisherGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublisherGetCommand_instances, "m", _PpSolutionPublisherGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving a specific publisher '${args.options.id || args.options.name}'...`);
        }
        const res = await this.getSolutionPublisher(args);
        await logger.log(res);
    }
    async getSolutionPublisher(args) {
        const requestOptions = {
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            if (args.options.id) {
                requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/publishers(${args.options.id})?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`;
                const result = await request.get(requestOptions);
                return result;
            }
            requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/publishers?$filter=friendlyname eq '${args.options.name}'&$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix&api-version=9.1`;
            const result = await request.get(requestOptions);
            if (result.value.length === 0) {
                throw `The specified publisher '${args.options.name}' does not exist.`;
            }
            return result.value[0];
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpSolutionPublisherGetCommand_instances = new WeakSet(), _PpSolutionPublisherGetCommand_initTelemetry = function _PpSolutionPublisherGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpSolutionPublisherGetCommand_initOptions = function _PpSolutionPublisherGetCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--asAdmin'
    });
}, _PpSolutionPublisherGetCommand_initOptionSets = function _PpSolutionPublisherGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
}, _PpSolutionPublisherGetCommand_initValidators = function _PpSolutionPublisherGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpSolutionPublisherGetCommand();
//# sourceMappingURL=solution-publisher-get.js.map