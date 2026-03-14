var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowRunCancelCommand_instances, _FlowRunCancelCommand_initTelemetry, _FlowRunCancelCommand_initOptions, _FlowRunCancelCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
class FlowRunCancelCommand extends PowerAutomateCommand {
    get name() {
        return commands.RUN_CANCEL;
    }
    get description() {
        return 'Cancels a specific run of the specified Microsoft Flow';
    }
    constructor() {
        super();
        _FlowRunCancelCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowRunCancelCommand_instances, "m", _FlowRunCancelCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowRunCancelCommand_instances, "m", _FlowRunCancelCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowRunCancelCommand_instances, "m", _FlowRunCancelCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.log(`Cancelling run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
        }
        const cancelFlow = async () => {
            const requestOptions = {
                url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/runs/${formatting.encodeQueryParameter(args.options.name)}/cancel?api-version=2016-11-01`,
                headers: {
                    accept: 'application/json'
                },
                responseType: 'json'
            };
            try {
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await cancelFlow();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to cancel the flow run ${args.options.name}?` });
            if (result) {
                await cancelFlow();
            }
        }
    }
}
_FlowRunCancelCommand_instances = new WeakSet(), _FlowRunCancelCommand_initTelemetry = function _FlowRunCancelCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _FlowRunCancelCommand_initOptions = function _FlowRunCancelCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '--flowName <flowName>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '-f, --force'
    });
}, _FlowRunCancelCommand_initValidators = function _FlowRunCancelCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.flowName)) {
            return `${args.options.flowName} is not a valid GUID`;
        }
        return true;
    });
};
export default new FlowRunCancelCommand();
//# sourceMappingURL=run-cancel.js.map