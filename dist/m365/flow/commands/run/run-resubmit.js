var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowRunResubmitCommand_instances, _FlowRunResubmitCommand_initTelemetry, _FlowRunResubmitCommand_initOptions, _FlowRunResubmitCommand_initValidators;
import chalk from 'chalk';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
class FlowRunResubmitCommand extends PowerAutomateCommand {
    get name() {
        return commands.RUN_RESUBMIT;
    }
    get description() {
        return 'Resubmits a specific flow run for the specified Microsoft Flow';
    }
    constructor() {
        super();
        _FlowRunResubmitCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowRunResubmitCommand_instances, "m", _FlowRunResubmitCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowRunResubmitCommand_instances, "m", _FlowRunResubmitCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowRunResubmitCommand_instances, "m", _FlowRunResubmitCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Resubmitting run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
        }
        const resubmitFlow = async () => {
            try {
                const triggerName = await this.getTriggerName(args.options.environmentName, args.options.flowName);
                if (this.debug) {
                    await logger.logToStderr(chalk.yellow(`Retrieved trigger: ${triggerName}`));
                }
                const requestOptions = {
                    url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/triggers/${formatting.encodeQueryParameter(triggerName)}/histories/${formatting.encodeQueryParameter(args.options.name)}/resubmit?api-version=2016-11-01`,
                    headers: {
                        accept: 'application/json'
                    },
                    responseType: 'json'
                };
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await resubmitFlow();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to resubmit the flow with run ${args.options.name}?` });
            if (result) {
                await resubmitFlow();
            }
        }
    }
    async getTriggerName(environment, flow) {
        const requestOptions = {
            url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(environment)}/flows/${formatting.encodeQueryParameter(flow)}/triggers?api-version=2016-11-01`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        const res = await request.get(requestOptions);
        return res.value[0].name;
    }
}
_FlowRunResubmitCommand_instances = new WeakSet(), _FlowRunResubmitCommand_initTelemetry = function _FlowRunResubmitCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: args.options.force
        });
    });
}, _FlowRunResubmitCommand_initOptions = function _FlowRunResubmitCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '--flowName <flowName>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-f, --force'
    });
}, _FlowRunResubmitCommand_initValidators = function _FlowRunResubmitCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.flowName)) {
            return `${args.options.flowName} is not a valid GUID`;
        }
        return true;
    });
};
export default new FlowRunResubmitCommand();
//# sourceMappingURL=run-resubmit.js.map