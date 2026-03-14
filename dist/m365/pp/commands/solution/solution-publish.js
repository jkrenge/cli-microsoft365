var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpSolutionPublishCommand_instances, _PpSolutionPublishCommand_initTelemetry, _PpSolutionPublishCommand_initOptions, _PpSolutionPublishCommand_initOptionSets, _PpSolutionPublishCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpSolutionPublishCommand extends PowerPlatformCommand {
    get name() {
        return commands.SOLUTION_PUBLISH;
    }
    get description() {
        return 'Publishes the components of a solution in a given Power Platform environment';
    }
    constructor() {
        super();
        _PpSolutionPublishCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpSolutionPublishCommand_instances, "m", _PpSolutionPublishCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublishCommand_instances, "m", _PpSolutionPublishCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublishCommand_instances, "m", _PpSolutionPublishCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpSolutionPublishCommand_instances, "m", _PpSolutionPublishCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);
            const solutionId = await this.getSolutionId(args, dynamicsApiUrl, logger);
            const solutionComponents = await this.getSolutionComponents(dynamicsApiUrl, solutionId, logger);
            const parameterXml = await this.buildXmlRequestObject(solutionComponents, logger);
            const requestOptions = {
                url: `${dynamicsApiUrl}/api/data/v9.0/PublishXml`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    ParameterXml: parameterXml
                }
            };
            if (this.verbose) {
                await logger.logToStderr(`Publishing the solution '${args.options.id || args.options.name}'...`);
            }
            if (args.options.wait) {
                await request.post(requestOptions);
            }
            else {
                void request.post(requestOptions);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getSolutionComponents(dynamicsApiUrl, solutionId, logger) {
        const requestOptions = {
            url: `${dynamicsApiUrl}/api/data/v9.0/msdyn_solutioncomponentsummaries?$filter=(msdyn_solutionid eq ${solutionId})&$select=msdyn_componentlogicalname,msdyn_name&$orderby=msdyn_componentlogicalname asc&api-version=9.1`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        if (this.verbose) {
            await logger.logToStderr(`Retrieving solution components`);
        }
        const response = await request.get(requestOptions);
        return response.value;
    }
    async getSolutionId(args, dynamicsApiUrl, logger) {
        if (args.options.id) {
            return args.options.id;
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving solutionId...`);
        }
        const solution = await powerPlatform.getSolutionByName(dynamicsApiUrl, args.options.name);
        return solution.solutionid;
    }
    async buildXmlRequestObject(solutionComponents, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Building the XML request object...`);
        }
        const result = solutionComponents.reduce(function (r, a) {
            const key = a.msdyn_componentlogicalname.slice(-1) === 'y' ?
                a.msdyn_componentlogicalname.substring(0, a.msdyn_componentlogicalname.length - 1) + 'ies' :
                a.msdyn_componentlogicalname + 's';
            r[key] = r[key] || [];
            r[key].push({ [a.msdyn_componentlogicalname]: a.msdyn_name });
            return r;
        }, Object.create(null));
        return `<importexportxml>${formatting.objectToXml(result)}</importexportxml>`;
    }
}
_PpSolutionPublishCommand_instances = new WeakSet(), _PpSolutionPublishCommand_initTelemetry = function _PpSolutionPublishCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            asAdmin: !!args.options.asAdmin,
            wait: !!args.options.wait
        });
    });
}, _PpSolutionPublishCommand_initOptions = function _PpSolutionPublishCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--asAdmin'
    }, {
        option: '--wait'
    });
}, _PpSolutionPublishCommand_initOptionSets = function _PpSolutionPublishCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
}, _PpSolutionPublishCommand_initValidators = function _PpSolutionPublishCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new PpSolutionPublishCommand();
//# sourceMappingURL=solution-publish.js.map