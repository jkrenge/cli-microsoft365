var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowRunGetCommand_instances, _FlowRunGetCommand_initOptions, _FlowRunGetCommand_initValidators, _FlowRunGetCommand_initTelemetry;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
class FlowRunGetCommand extends PowerAutomateCommand {
    get name() {
        return commands.RUN_GET;
    }
    get description() {
        return 'Gets information about a specific run of the specified Microsoft Flow';
    }
    constructor() {
        super();
        _FlowRunGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowRunGetCommand_instances, "m", _FlowRunGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowRunGetCommand_instances, "m", _FlowRunGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowRunGetCommand_instances, "m", _FlowRunGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
        }
        const actionsParameter = args.options.withActions ? '$expand=properties%2Factions&' : '';
        const requestOptions = {
            url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/runs/${formatting.encodeQueryParameter(args.options.name)}?${actionsParameter}api-version=2016-11-01`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            res.startTime = res.properties.startTime;
            res.endTime = res.properties.endTime || '';
            res.status = res.properties.status;
            res.triggerName = res.properties.trigger.name;
            if (args.options.withTrigger && res.properties.trigger.outputsLink) {
                res.triggerInformation = await this.getTriggerInformation(res);
            }
            if (args.options.withActions) {
                res.actions = await this.getActionsInformation(res, args.options.withActions);
            }
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTriggerInformation(res) {
        return await this.requestAdditionalInformation(res.properties.trigger.outputsLink.uri);
    }
    async getActionsInformation(res, withActions) {
        const chosenActions = typeof withActions === 'string' ? withActions.split(',') : null;
        const actionsResult = {};
        for (const action in res.properties.actions) {
            if (!res.properties.actions[action] || (chosenActions && chosenActions.indexOf(action) === -1)) {
                continue;
            }
            actionsResult[action] = res.properties.actions[action];
            if (res.properties.actions[action].inputsLink?.uri) {
                actionsResult[action].input = await this.requestAdditionalInformation(res.properties.actions[action].inputsLink?.uri);
            }
            if (res.properties.actions[action].outputsLink?.uri) {
                actionsResult[action].output = await this.requestAdditionalInformation(res.properties.actions[action].outputsLink?.uri);
            }
        }
        return actionsResult;
    }
    async requestAdditionalInformation(requestUri) {
        const additionalInformationOptions = {
            url: requestUri,
            headers: {
                accept: 'application/json',
                'x-anonymous': true
            },
            responseType: 'json'
        };
        const additionalInformationResponse = await request.get(additionalInformationOptions);
        return additionalInformationResponse.body ? additionalInformationResponse.body : additionalInformationResponse;
    }
}
_FlowRunGetCommand_instances = new WeakSet(), _FlowRunGetCommand_initOptions = function _FlowRunGetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '--flowName <flowName>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--withTrigger'
    }, {
        option: '--withActions [withActions]'
    });
}, _FlowRunGetCommand_initValidators = function _FlowRunGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.flowName)) {
            return `${args.options.flowName} is not a valid GUID`;
        }
        if (args.options.withActions && (typeof args.options.withActions !== 'string' && typeof args.options.withActions !== 'boolean')) {
            return 'the withActions parameter must be a string or boolean';
        }
        return true;
    });
}, _FlowRunGetCommand_initTelemetry = function _FlowRunGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            withTrigger: !!args.options.withTrigger,
            withActions: typeof args.options.withActions !== 'undefined'
        });
    });
};
export default new FlowRunGetCommand();
//# sourceMappingURL=run-get.js.map