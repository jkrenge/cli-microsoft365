var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpSolutionGetCommand_instances, _PpSolutionGetCommand_initTelemetry, _PpSolutionGetCommand_initOptions, _PpSolutionGetCommand_initOptionSets, _PpSolutionGetCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpSolutionGetCommand extends PowerPlatformCommand {
    get name() {
        return commands.SOLUTION_GET;
    }
    get description() {
        return 'Gets a specific solution in a given environment.';
    }
    constructor() {
        super();
        _PpSolutionGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpSolutionGetCommand_instances, "m", _PpSolutionGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpSolutionGetCommand_instances, "m", _PpSolutionGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpSolutionGetCommand_instances, "m", _PpSolutionGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpSolutionGetCommand_instances, "m", _PpSolutionGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving a specific solution '${args.options.id || args.options.name}'...`);
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const res = await this.getSolution(dynamicsApiUrl, args.options);
            if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
                await logger.log(res);
            }
            else {
                // Converted to text friendly output
                await logger.log({
                    uniquename: res.uniquename,
                    version: res.version,
                    publisher: res.publisherid.friendlyname
                });
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getSolution(dynamicsApiUrl, options) {
        if (options.name) {
            return powerPlatform.getSolutionByName(dynamicsApiUrl, options.name);
        }
        const requestOptions = {
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        requestOptions.url = `${dynamicsApiUrl}/api/data/v9.0/solutions(${options.id})?$expand=publisherid($select=friendlyname)&$select=solutionid,uniquename,version,publisherid,installedon,solutionpackageversion,friendlyname,versionnumber&api-version=9.1`;
        return request.get(requestOptions);
    }
}
_PpSolutionGetCommand_instances = new WeakSet(), _PpSolutionGetCommand_initTelemetry = function _PpSolutionGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpSolutionGetCommand_initOptions = function _PpSolutionGetCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--asAdmin'
    });
}, _PpSolutionGetCommand_initOptionSets = function _PpSolutionGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
}, _PpSolutionGetCommand_initValidators = function _PpSolutionGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpSolutionGetCommand();
//# sourceMappingURL=solution-get.js.map