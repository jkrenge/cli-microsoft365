var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserRecycleBinItemRestoreCommand_instances, _EntraUserRecycleBinItemRestoreCommand_initOptions, _EntraUserRecycleBinItemRestoreCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraUserRecycleBinItemRestoreCommand extends GraphCommand {
    get name() {
        return commands.USER_RECYCLEBINITEM_RESTORE;
    }
    get description() {
        return 'Restores a user from the tenant recycle bin';
    }
    constructor() {
        super();
        _EntraUserRecycleBinItemRestoreCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserRecycleBinItemRestoreCommand_instances, "m", _EntraUserRecycleBinItemRestoreCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserRecycleBinItemRestoreCommand_instances, "m", _EntraUserRecycleBinItemRestoreCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Restoring user with id ${args.options.id} from the recycle bin.`);
        }
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/directory/deletedItems/${args.options.id}/restore`,
                headers: {
                    'content-type': 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const user = await request.post(requestOptions);
            await logger.log(user);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraUserRecycleBinItemRestoreCommand_instances = new WeakSet(), _EntraUserRecycleBinItemRestoreCommand_initOptions = function _EntraUserRecycleBinItemRestoreCommand_initOptions() {
    this.options.unshift({
        option: '--id <id>'
    });
}, _EntraUserRecycleBinItemRestoreCommand_initValidators = function _EntraUserRecycleBinItemRestoreCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new EntraUserRecycleBinItemRestoreCommand();
//# sourceMappingURL=user-recyclebinitem-restore.js.map