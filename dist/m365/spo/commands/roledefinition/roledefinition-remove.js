var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoRoleDefinitionRemoveCommand_instances, _SpoRoleDefinitionRemoveCommand_initTelemetry, _SpoRoleDefinitionRemoveCommand_initOptions, _SpoRoleDefinitionRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoRoleDefinitionRemoveCommand extends SpoCommand {
    get name() {
        return commands.ROLEDEFINITION_REMOVE;
    }
    get description() {
        return 'Removes the role definition from the specified site';
    }
    constructor() {
        super();
        _SpoRoleDefinitionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionRemoveCommand_instances, "m", _SpoRoleDefinitionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionRemoveCommand_instances, "m", _SpoRoleDefinitionRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoRoleDefinitionRemoveCommand_instances, "m", _SpoRoleDefinitionRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeRoleDefinition(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the role definition with id ${args.options.id} from site ${args.options.webUrl}?` });
            if (result) {
                await this.removeRoleDefinition(logger, args);
            }
        }
    }
    async removeRoleDefinition(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing role definition from site ${args.options.webUrl}...`);
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/roledefinitions(${args.options.id})`,
            method: 'delete',
            headers: {
                'X-HTTP-Method': 'DELETE',
                'If-Match': '*',
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            await request.delete(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoRoleDefinitionRemoveCommand_instances = new WeakSet(), _SpoRoleDefinitionRemoveCommand_initTelemetry = function _SpoRoleDefinitionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoRoleDefinitionRemoveCommand_initOptions = function _SpoRoleDefinitionRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _SpoRoleDefinitionRemoveCommand_initValidators = function _SpoRoleDefinitionRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const id = parseInt(args.options.id);
        if (isNaN(id)) {
            return `${args.options.id} is not a valid role definition ID`;
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoRoleDefinitionRemoveCommand();
//# sourceMappingURL=roledefinition-remove.js.map