var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionEventTypeAddCommand_instances, _PurviewRetentionEventTypeAddCommand_initTelemetry, _PurviewRetentionEventTypeAddCommand_initOptions;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionEventTypeAddCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONEVENTTYPE_ADD;
    }
    get description() {
        return 'Create a retention event type';
    }
    constructor() {
        super();
        _PurviewRetentionEventTypeAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeAddCommand_instances, "m", _PurviewRetentionEventTypeAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeAddCommand_instances, "m", _PurviewRetentionEventTypeAddCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const requestBody = {
            displayName: args.options.displayName,
            description: args.options.description
        };
        const requestOptions = {
            url: `${this.resource}/v1.0/security/triggerTypes/retentionEventTypes`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            data: requestBody,
            responseType: 'json'
        };
        try {
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataPromise(err);
        }
    }
}
_PurviewRetentionEventTypeAddCommand_instances = new WeakSet(), _PurviewRetentionEventTypeAddCommand_initTelemetry = function _PurviewRetentionEventTypeAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined'
        });
    });
}, _PurviewRetentionEventTypeAddCommand_initOptions = function _PurviewRetentionEventTypeAddCommand_initOptions() {
    this.options.unshift({
        option: '-n, --displayName <displayName>'
    }, {
        option: '-d, --description [description]'
    });
};
export default new PurviewRetentionEventTypeAddCommand();
//# sourceMappingURL=retentioneventtype-add.js.map