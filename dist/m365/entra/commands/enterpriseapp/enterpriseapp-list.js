var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraEnterpriseAppListCommand_instances, _EntraEnterpriseAppListCommand_initTelemetry, _EntraEnterpriseAppListCommand_initOptions;
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraEnterpriseAppListCommand extends GraphCommand {
    get name() {
        return commands.ENTERPRISEAPP_LIST;
    }
    defaultProperties() {
        return ['appId', 'displayName', 'tag'];
    }
    get description() {
        return 'Lists the enterprise applications (or service principals) in Entra ID';
    }
    alias() {
        return [commands.SP_LIST];
    }
    constructor() {
        super();
        _EntraEnterpriseAppListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppListCommand_instances, "m", _EntraEnterpriseAppListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraEnterpriseAppListCommand_instances, "m", _EntraEnterpriseAppListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving enterprise application information...`);
        }
        try {
            let requestUrl = `${this.resource}/v1.0/servicePrincipals`;
            const filter = [];
            if (args.options.tag) {
                filter.push(`(tags/any(t:t eq '${args.options.tag}'))`);
            }
            if (args.options.displayName) {
                filter.push(`(displayName eq '${args.options.displayName}')`);
            }
            if (filter.length > 0) {
                requestUrl += `?$filter=${filter.join(' and ')}`;
            }
            const res = await odata.getAllItems(requestUrl);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraEnterpriseAppListCommand_instances = new WeakSet(), _EntraEnterpriseAppListCommand_initTelemetry = function _EntraEnterpriseAppListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            displayName: typeof args.options.displayName !== 'undefined',
            tag: typeof args.options.tag !== 'undefined'
        });
    });
}, _EntraEnterpriseAppListCommand_initOptions = function _EntraEnterpriseAppListCommand_initOptions() {
    this.options.unshift({
        option: '-n, --displayName [displayName]'
    }, {
        option: '--tag [tag]'
    });
};
export default new EntraEnterpriseAppListCommand();
//# sourceMappingURL=enterpriseapp-list.js.map