var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowEnableCommand_instances, _FlowEnableCommand_initTelemetry, _FlowEnableCommand_initOptions;
import request from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
import commands from '../commands.js';
class FlowEnableCommand extends PowerAutomateCommand {
    get name() {
        return commands.ENABLE;
    }
    get description() {
        return 'Enables specified Microsoft Flow';
    }
    constructor() {
        super();
        _FlowEnableCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowEnableCommand_instances, "m", _FlowEnableCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowEnableCommand_instances, "m", _FlowEnableCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Enables Microsoft Flow ${args.options.name}...`);
        }
        const requestOptions = {
            url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}/start?api-version=2016-11-01`,
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
_FlowEnableCommand_instances = new WeakSet(), _FlowEnableCommand_initTelemetry = function _FlowEnableCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _FlowEnableCommand_initOptions = function _FlowEnableCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--asAdmin'
    });
};
export default new FlowEnableCommand();
//# sourceMappingURL=flow-enable.js.map