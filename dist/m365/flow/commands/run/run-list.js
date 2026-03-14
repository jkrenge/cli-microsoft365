var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowRunListCommand_instances, _FlowRunListCommand_initTelemetry, _FlowRunListCommand_initOptions, _FlowRunListCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
class FlowRunListCommand extends PowerAutomateCommand {
    get name() {
        return commands.RUN_LIST;
    }
    get description() {
        return 'Lists runs of the specified Microsoft Flow';
    }
    defaultProperties() {
        return ['name', 'startTime', 'status'];
    }
    constructor() {
        super();
        _FlowRunListCommand_instances.add(this);
        this.allowedStatusses = ['Succeeded', 'Running', 'Failed', 'Cancelled'];
        __classPrivateFieldGet(this, _FlowRunListCommand_instances, "m", _FlowRunListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowRunListCommand_instances, "m", _FlowRunListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowRunListCommand_instances, "m", _FlowRunListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list of runs for Microsoft Flow ${args.options.flowName}...`);
        }
        let url = `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/runs?api-version=2016-11-01`;
        const filters = this.getFilters(args.options);
        if (filters.length > 0) {
            url += `&$filter=${filters.join(' and ')}`;
        }
        try {
            const items = await odata.getAllItems(url);
            if (args.options.withTrigger) {
                await this.retrieveTriggerInformation(items);
            }
            if (args.options.output !== 'json' && items.length > 0) {
                items.forEach(i => {
                    i.startTime = i.properties.startTime;
                    i.status = i.properties.status;
                });
            }
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getFilters(options) {
        const filters = [];
        if (options.status) {
            filters.push(`status eq '${options.status}'`);
        }
        if (options.triggerStartTime) {
            filters.push(`startTime ge ${options.triggerStartTime}`);
        }
        if (options.triggerEndTime) {
            filters.push(`startTime lt ${options.triggerEndTime}`);
        }
        return filters;
    }
    async retrieveTriggerInformation(items) {
        const tasks = items.map(async (item) => {
            const requestOptions = {
                url: item.properties.trigger.outputsLink.uri,
                headers: {
                    accept: 'application/json',
                    'x-anonymous': true
                },
                responseType: 'json'
            };
            const response = await request.get(requestOptions);
            item.triggerInformation = response.body;
        });
        await Promise.all(tasks);
    }
}
_FlowRunListCommand_instances = new WeakSet(), _FlowRunListCommand_initTelemetry = function _FlowRunListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            status: typeof args.options.status !== 'undefined',
            triggerStartTime: typeof args.options.triggerStartTime !== 'undefined',
            triggerEndTime: typeof args.options.triggerEndTime !== 'undefined',
            withTrigger: !!args.options.withTrigger,
            asAdmin: !!args.options.asAdmin
        });
    });
}, _FlowRunListCommand_initOptions = function _FlowRunListCommand_initOptions() {
    this.options.unshift({
        option: '--flowName <flowName>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--status [status]',
        autocomplete: this.allowedStatusses
    }, {
        option: '--triggerStartTime [triggerStartTime]'
    }, {
        option: '--triggerEndTime [triggerEndTime]'
    }, {
        option: '--withTrigger'
    }, {
        option: '--asAdmin'
    });
}, _FlowRunListCommand_initValidators = function _FlowRunListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.flowName)) {
            return `${args.options.flowName} is not a valid GUID`;
        }
        if (args.options.status && this.allowedStatusses.indexOf(args.options.status) === -1) {
            return `'${args.options.status}' is not a valid status. Allowed values are: ${this.allowedStatusses.join(',')}`;
        }
        if (args.options.triggerStartTime && !validation.isValidISODateTime(args.options.triggerStartTime)) {
            return `'${args.options.triggerStartTime}' is not a valid datetime.`;
        }
        if (args.options.triggerEndTime && !validation.isValidISODateTime(args.options.triggerEndTime)) {
            return `'${args.options.triggerEndTime}' is not a valid datetime.`;
        }
        if (args.options.output !== 'json' && args.options.withTrigger) {
            return 'The --withTrigger option is only available when output is set to json';
        }
        return true;
    });
};
export default new FlowRunListCommand();
//# sourceMappingURL=run-list.js.map