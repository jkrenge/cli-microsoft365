var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowDisableCommand_instances, _FlowDisableCommand_initTelemetry, _FlowDisableCommand_initOptions;
import request from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
import commands from '../commands.js';
class FlowDisableCommand extends PowerAutomateCommand {
    get name() {
        return commands.DISABLE;
    }
    get description() {
        return 'Disables specified Microsoft Flow';
    }
    constructor() {
        super();
        _FlowDisableCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowDisableCommand_instances, "m", _FlowDisableCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowDisableCommand_instances, "m", _FlowDisableCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Disables Microsoft Flow ${args.options.name}...`);
        }
        const requestOptions = {
            url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}/stop?api-version=2016-11-01`,
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
    }
}
_FlowDisableCommand_instances = new WeakSet(), _FlowDisableCommand_initTelemetry = function _FlowDisableCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _FlowDisableCommand_initOptions = function _FlowDisableCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--asAdmin'
    });
};
export default new FlowDisableCommand();
//# sourceMappingURL=flow-disable.js.map