var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _VivaEngageNetworkListCommand_instances, _VivaEngageNetworkListCommand_initTelemetry, _VivaEngageNetworkListCommand_initOptions;
import request from '../../../../request.js';
import VivaEngageCommand from '../../../base/VivaEngageCommand.js';
import commands from '../../commands.js';
class VivaEngageNetworkListCommand extends VivaEngageCommand {
    get name() {
        return commands.ENGAGE_NETWORK_LIST;
    }
    get description() {
        return 'Returns a list of networks to which the current user has access';
    }
    defaultProperties() {
        return ['id', 'name', 'email', 'community', 'permalink', 'web_url'];
    }
    constructor() {
        super();
        _VivaEngageNetworkListCommand_instances.add(this);
        __classPrivateFieldGet(this, _VivaEngageNetworkListCommand_instances, "m", _VivaEngageNetworkListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _VivaEngageNetworkListCommand_instances, "m", _VivaEngageNetworkListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/v1/networks/current.json`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
            },
            responseType: 'json',
            data: {
                includeSuspended: (args.options.withSuspended !== undefined && args.options.withSuspended !== false)
            }
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_VivaEngageNetworkListCommand_instances = new WeakSet(), _VivaEngageNetworkListCommand_initTelemetry = function _VivaEngageNetworkListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            withSuspended: args.options.withSuspended
        });
    });
}, _VivaEngageNetworkListCommand_initOptions = function _VivaEngageNetworkListCommand_initOptions() {
    this.options.unshift({
        option: '--withSuspended'
    });
};
export default new VivaEngageNetworkListCommand();
//# sourceMappingURL=engage-network-list.js.map