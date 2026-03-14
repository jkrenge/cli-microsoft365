var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsAppListCommand_instances, _a, _TeamsAppListCommand_initTelemetry, _TeamsAppListCommand_initOptions, _TeamsAppListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsAppListCommand extends GraphCommand {
    get name() {
        return commands.APP_LIST;
    }
    get description() {
        return 'Lists apps from the Microsoft Teams app catalog';
    }
    defaultProperties() {
        return ['id', 'displayName', 'distributionMethod'];
    }
    constructor() {
        super();
        _TeamsAppListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsAppListCommand_instances, "m", _TeamsAppListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsAppListCommand_instances, "m", _TeamsAppListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsAppListCommand_instances, "m", _TeamsAppListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            let requestUrl = `${this.resource}/v1.0/appCatalogs/teamsApps`;
            if (args.options.distributionMethod) {
                requestUrl += `?$filter=distributionMethod eq '${args.options.distributionMethod}'`;
            }
            const items = await odata.getAllItems(requestUrl);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = TeamsAppListCommand, _TeamsAppListCommand_instances = new WeakSet(), _TeamsAppListCommand_initTelemetry = function _TeamsAppListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            distributionMethod: args.options.distributionMethod || false
        });
    });
}, _TeamsAppListCommand_initOptions = function _TeamsAppListCommand_initOptions() {
    this.options.unshift({
        option: '--distributionMethod [distributionMethod]',
        autocomplete: _a.allowedDistributionMethods
    });
}, _TeamsAppListCommand_initValidators = function _TeamsAppListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.distributionMethod &&
            _a.allowedDistributionMethods.indexOf(args.options.distributionMethod) < 0) {
            return `'${args.options.distributionMethod}' is not a valid distributionMethod. Allowed distribution methods are: ${_a.allowedDistributionMethods.join(', ')}`;
        }
        return true;
    });
};
TeamsAppListCommand.allowedDistributionMethods = ['store', 'organization', 'sideloaded'];
export default new TeamsAppListCommand();
//# sourceMappingURL=app-list.js.map