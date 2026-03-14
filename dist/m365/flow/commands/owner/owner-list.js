var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _FlowOwnerListCommand_instances, _FlowOwnerListCommand_initTelemetry, _FlowOwnerListCommand_initOptions, _FlowOwnerListCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';
class FlowOwnerListCommand extends PowerAutomateCommand {
    get name() {
        return commands.OWNER_LIST;
    }
    get description() {
        return 'Lists all owners of a Power Automate flow';
    }
    defaultProperties() {
        return ['roleName', 'id', 'type'];
    }
    constructor() {
        super();
        _FlowOwnerListCommand_instances.add(this);
        __classPrivateFieldGet(this, _FlowOwnerListCommand_instances, "m", _FlowOwnerListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _FlowOwnerListCommand_instances, "m", _FlowOwnerListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _FlowOwnerListCommand_instances, "m", _FlowOwnerListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Listing owners for flow ${args.options.flowName} in environment ${args.options.environmentName}`);
            }
            const response = await odata.getAllItems(`${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/permissions?api-version=2016-11-01`);
            if (!cli.shouldTrimOutput(args.options.output)) {
                await logger.log(response);
            }
            else {
                //converted to text friendly output
                await logger.log(response.map(res => ({
                    roleName: res.properties.roleName,
                    id: res.properties.principal.id,
                    type: res.properties.principal.type
                })));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_FlowOwnerListCommand_instances = new WeakSet(), _FlowOwnerListCommand_initTelemetry = function _FlowOwnerListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _FlowOwnerListCommand_initOptions = function _FlowOwnerListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--flowName <flowName>'
    }, {
        option: '--asAdmin'
    });
}, _FlowOwnerListCommand_initValidators = function _FlowOwnerListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.flowName)) {
            return `${args.options.flowName} is not a valid GUID.`;
        }
        return true;
    });
};
export default new FlowOwnerListCommand();
//# sourceMappingURL=owner-list.js.map