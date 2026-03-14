var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionEventGetCommand_instances, _PurviewRetentionEventGetCommand_initOptions, _PurviewRetentionEventGetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionEventGetCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONEVENT_GET;
    }
    get description() {
        return 'Retrieve the specified retention event';
    }
    constructor() {
        super();
        _PurviewRetentionEventGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventGetCommand_instances, "m", _PurviewRetentionEventGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionEventGetCommand_instances, "m", _PurviewRetentionEventGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving retention event with id ${args.options.id}`);
        }
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/security/triggers/retentionEvents/${args.options.id}`,
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
_PurviewRetentionEventGetCommand_instances = new WeakSet(), _PurviewRetentionEventGetCommand_initOptions = function _PurviewRetentionEventGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
}, _PurviewRetentionEventGetCommand_initValidators = function _PurviewRetentionEventGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID.`;
        }
        return true;
    });
};
export default new PurviewRetentionEventGetCommand();
//# sourceMappingURL=retentionevent-get.js.map