var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowGetCommand_instances, _FlowGetCommand_initTelemetry, _FlowGetCommand_initOptions;
import request from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import PowerAutomateCommand from '../../base/PowerAutomateCommand.js';
import commands from '../commands.js';
class FlowGetCommand extends PowerAutomateCommand {
    get name() {
        return commands.GET;
    }
    get description() {
        return 'Gets information about the specified Microsoft Flow';
    }
    constructor() {
        super();
        _FlowGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowGetCommand_instances, "m", _FlowGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowGetCommand_instances, "m", _FlowGetCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about Microsoft Flow ${args.options.name}...`);
        }
        const requestOptions = {
            url: `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01&$expand=swagger,properties.connectionreferences.apidefinition,properties.definitionsummary.operations.apioperation,operationDefinition,plan,properties.throttleData,properties.estimatedsuspensiondata,properties.licenseData`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            res.displayName = res.properties.displayName;
            res.description = res.properties.definitionSummary.description || '';
            res.triggers = res.properties.definitionSummary.triggers.map((t) => (t.type + (t.kind ? "-" + t.kind : ''))).join(', ');
            res.actions = res.properties.definitionSummary.actions.map((a) => (a.type + (a.swaggerOperationId ? "-" + a.swaggerOperationId : ''))).join(', ');
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_FlowGetCommand_instances = new WeakSet(), _FlowGetCommand_initTelemetry = function _FlowGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _FlowGetCommand_initOptions = function _FlowGetCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>'
    }, {
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--asAdmin'
    });
};
export default new FlowGetCommand();
//# sourceMappingURL=flow-get.js.map