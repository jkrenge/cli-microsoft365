var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupRemoveCommand_instances, _EntraGroupRemoveCommand_initTelemetry, _EntraGroupRemoveCommand_initOptions, _EntraGroupRemoveCommand_initOptionSets, _EntraGroupRemoveCommand_initValidators, _EntraGroupRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { validation } from '../../../../utils/validation.js';
class EntraGroupRemoveCommand extends GraphCommand {
    get name() {
        return commands.GROUP_REMOVE;
    }
    get description() {
        return 'Removes an Entra group';
    }
    constructor() {
        super();
        _EntraGroupRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupRemoveCommand_instances, "m", _EntraGroupRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupRemoveCommand_instances, "m", _EntraGroupRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraGroupRemoveCommand_instances, "m", _EntraGroupRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraGroupRemoveCommand_instances, "m", _EntraGroupRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupRemoveCommand_instances, "m", _EntraGroupRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeGroup = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing group ${args.options.id || args.options.displayName}...`);
            }
            try {
                let groupId = args.options.id;
                if (args.options.displayName) {
                    groupId = await entraGroup.getGroupIdByDisplayName(args.options.displayName);
                }
                const requestOptions = {
                    url: `${this.resource}/v1.0/groups/${groupId}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    }
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeGroup();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove group '${args.options.id || args.options.displayName}'?` });
            if (result) {
                await removeGroup();
            }
        }
    }
}
_EntraGroupRemoveCommand_instances = new WeakSet(), _EntraGroupRemoveCommand_initTelemetry = function _EntraGroupRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: args.options.id !== 'undefined',
            displayName: args.options.displayName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _EntraGroupRemoveCommand_initOptions = function _EntraGroupRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    }, {
        option: '-f, --force'
    });
}, _EntraGroupRemoveCommand_initOptionSets = function _EntraGroupRemoveCommand_initOptionSets() {
    this.optionSets.push({
        options: ['id', 'displayName']
    });
}, _EntraGroupRemoveCommand_initValidators = function _EntraGroupRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID for option id.`;
        }
        return true;
    });
}, _EntraGroupRemoveCommand_initTypes = function _EntraGroupRemoveCommand_initTypes() {
    this.types.string.push('id', 'displayName');
};
export default new EntraGroupRemoveCommand();
//# sourceMappingURL=group-remove.js.map