var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppListCommand_instances, _EntraAppListCommand_initTelemetry, _EntraAppListCommand_initOptions;
import { odata } from "../../../../utils/odata.js";
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraAppListCommand extends GraphCommand {
    get name() {
        return commands.APP_LIST;
    }
    get description() {
        return 'Retrieves a list of Entra app registrations';
    }
    constructor() {
        super();
        _EntraAppListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppListCommand_instances, "m", _EntraAppListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppListCommand_instances, "m", _EntraAppListCommand_initOptions).call(this);
    }
    defaultProperties() {
        return ['appId', 'id', 'displayName', "signInAudience"];
    }
    async commandAction(logger, args) {
        const queryParameters = [];
        if (args.options.properties) {
            const allProperties = args.options.properties.split(',');
            const selectProperties = allProperties.filter(prop => !prop.includes('/'));
            if (selectProperties.length > 0) {
                queryParameters.push(`$select=${selectProperties}`);
            }
        }
        const queryString = queryParameters.length > 0
            ? `?${queryParameters.join('&')}`
            : '';
        try {
            const results = await odata.getAllItems(`${this.resource}/v1.0/applications${queryString}`);
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraAppListCommand_instances = new WeakSet(), _EntraAppListCommand_initTelemetry = function _EntraAppListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            properties: typeof args.options.properties !== 'undefined'
        });
    });
}, _EntraAppListCommand_initOptions = function _EntraAppListCommand_initOptions() {
    this.options.unshift({ option: '-p, --properties [properties]' });
};
export default new EntraAppListCommand();
//# sourceMappingURL=app-list.js.map