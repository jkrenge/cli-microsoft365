var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionEventTypeGetCommand_instances, _PurviewRetentionEventTypeGetCommand_initOptions, _PurviewRetentionEventTypeGetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionEventTypeGetCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONEVENTTYPE_GET;
    }
    get description() {
        return 'Retrieve the specified retention event type';
    }
    constructor() {
        super();
        _PurviewRetentionEventTypeGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeGetCommand_instances, "m", _PurviewRetentionEventTypeGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventTypeGetCommand_instances, "m", _PurviewRetentionEventTypeGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving retention event type with id ${args.options.id}`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/security/triggerTypes/retentionEventTypes/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PurviewRetentionEventTypeGetCommand_instances = new WeakSet(), _PurviewRetentionEventTypeGetCommand_initOptions = function _PurviewRetentionEventTypeGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
}, _PurviewRetentionEventTypeGetCommand_initValidators = function _PurviewRetentionEventTypeGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID.`;
        }
        return true;
    });
};
export default new PurviewRetentionEventTypeGetCommand();
//# sourceMappingURL=retentioneventtype-get.js.map