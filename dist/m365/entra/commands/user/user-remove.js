var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserRemoveCommand_instances, _EntraUserRemoveCommand_initTelemetry, _EntraUserRemoveCommand_initOptions, _EntraUserRemoveCommand_initOptionSets, _EntraUserRemoveCommand_initValidators;
import commands from '../../commands.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { cli } from '../../../../cli/cli.js';
import GraphCommand from '../../../base/GraphCommand.js';
class EntraUserRemoveCommand extends GraphCommand {
    get name() {
        return commands.USER_REMOVE;
    }
    get description() {
        return 'Removes a specific user';
    }
    constructor() {
        super();
        _EntraUserRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserRemoveCommand_instances, "m", _EntraUserRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserRemoveCommand_instances, "m", _EntraUserRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserRemoveCommand_instances, "m", _EntraUserRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraUserRemoveCommand_instances, "m", _EntraUserRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing user '${args.options.id || args.options.userName}'...`);
        }
        if (args.options.force) {
            await this.deleteUser(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove user '${args.options.id || args.options.userName}'?` });
            if (result) {
                await this.deleteUser(args);
            }
        }
    }
    async deleteUser(args) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/users/${args.options.id || args.options.userName}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraUserRemoveCommand_instances = new WeakSet(), _EntraUserRemoveCommand_initTelemetry = function _EntraUserRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _EntraUserRemoveCommand_initOptions = function _EntraUserRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--id [id]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '-f, --force'
    });
}, _EntraUserRemoveCommand_initOptionSets = function _EntraUserRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'userName'] });
}, _EntraUserRemoveCommand_initValidators = function _EntraUserRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN)`;
        }
        return true;
    });
};
export default new EntraUserRemoveCommand();
//# sourceMappingURL=user-remove.js.map