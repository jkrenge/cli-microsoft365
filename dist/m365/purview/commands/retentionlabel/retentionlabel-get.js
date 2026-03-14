var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewRetentionLabelGetCommand_instances, _PurviewRetentionLabelGetCommand_initOptions, _PurviewRetentionLabelGetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PurviewRetentionLabelGetCommand extends GraphCommand {
    get name() {
        return commands.RETENTIONLABEL_GET;
    }
    get description() {
        return 'Retrieve the specified retention label';
    }
    constructor() {
        super();
        _PurviewRetentionLabelGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelGetCommand_instances, "m", _PurviewRetentionLabelGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewRetentionLabelGetCommand_instances, "m", _PurviewRetentionLabelGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving retention label with id ${args.options.id}`);
            }
            const requestOptions = {
                url: `${this.resource}/beta/security/labels/retentionLabels/${args.options.id}`,
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
_PurviewRetentionLabelGetCommand_instances = new WeakSet(), _PurviewRetentionLabelGetCommand_initOptions = function _PurviewRetentionLabelGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
}, _PurviewRetentionLabelGetCommand_initValidators = function _PurviewRetentionLabelGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID.`;
        }
        return true;
    });
};
export default new PurviewRetentionLabelGetCommand();
//# sourceMappingURL=retentionlabel-get.js.map