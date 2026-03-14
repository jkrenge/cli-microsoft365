var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserRecycleBinItemRemoveCommand_instances, _EntraUserRecycleBinItemRemoveCommand_initTelemetry, _EntraUserRecycleBinItemRemoveCommand_initOptions, _EntraUserRecycleBinItemRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraUserRecycleBinItemRemoveCommand extends GraphCommand {
    get name() {
        return commands.USER_RECYCLEBINITEM_REMOVE;
    }
    get description() {
        return 'Removes a user from the recycle bin in the current tenant';
    }
    constructor() {
        super();
        _EntraUserRecycleBinItemRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserRecycleBinItemRemoveCommand_instances, "m", _EntraUserRecycleBinItemRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserRecycleBinItemRemoveCommand_instances, "m", _EntraUserRecycleBinItemRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserRecycleBinItemRemoveCommand_instances, "m", _EntraUserRecycleBinItemRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const clearRecycleBinItem = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Permanently deleting user with id ${args.options.id} from Microsoft Entra ID`);
            }
            try {
                const requestOptions = {
                    url: `${this.resource}/v1.0/directory/deletedItems/${args.options.id}`,
                    headers: {}
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await clearRecycleBinItem();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to permanently delete the user with id ${args.options.id}?` });
            if (result) {
                await clearRecycleBinItem();
            }
        }
    }
}
_EntraUserRecycleBinItemRemoveCommand_instances = new WeakSet(), _EntraUserRecycleBinItemRemoveCommand_initTelemetry = function _EntraUserRecycleBinItemRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _EntraUserRecycleBinItemRemoveCommand_initOptions = function _EntraUserRecycleBinItemRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--id <id>'
    }, {
        option: '-f, --force'
    });
}, _EntraUserRecycleBinItemRemoveCommand_initValidators = function _EntraUserRecycleBinItemRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new EntraUserRecycleBinItemRemoveCommand();
//# sourceMappingURL=user-recyclebinitem-remove.js.map