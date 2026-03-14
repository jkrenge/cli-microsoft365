var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TenantPeopleProfileCardPropertyRemoveCommand_instances, _TenantPeopleProfileCardPropertyRemoveCommand_initTelemetry, _TenantPeopleProfileCardPropertyRemoveCommand_initOptions, _TenantPeopleProfileCardPropertyRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { profileCardPropertyNames } from './profileCardProperties.js';
class TenantPeopleProfileCardPropertyRemoveCommand extends GraphCommand {
    get name() {
        return commands.PEOPLE_PROFILECARDPROPERTY_REMOVE;
    }
    get description() {
        return 'Removes an additional attribute from the profile card properties';
    }
    constructor() {
        super();
        _TenantPeopleProfileCardPropertyRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertyRemoveCommand_instances, "m", _TenantPeopleProfileCardPropertyRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertyRemoveCommand_instances, "m", _TenantPeopleProfileCardPropertyRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TenantPeopleProfileCardPropertyRemoveCommand_instances, "m", _TenantPeopleProfileCardPropertyRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const directoryPropertyName = profileCardPropertyNames.find(n => n.toLowerCase() === args.options.name.toLowerCase());
        const removeProfileCardProperty = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing '${directoryPropertyName}' as a profile card property...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/admin/people/profileCardProperties/${directoryPropertyName}`,
                headers: {
                    'content-type': 'application/json'
                },
                responseType: 'json'
            };
            try {
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeProfileCardProperty();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the profile card property '${directoryPropertyName}'?` });
            if (result) {
                await removeProfileCardProperty();
            }
        }
    }
}
_TenantPeopleProfileCardPropertyRemoveCommand_instances = new WeakSet(), _TenantPeopleProfileCardPropertyRemoveCommand_initTelemetry = function _TenantPeopleProfileCardPropertyRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: !!args.options.force
        });
    });
}, _TenantPeopleProfileCardPropertyRemoveCommand_initOptions = function _TenantPeopleProfileCardPropertyRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-n, --name <name>',
        autocomplete: profileCardPropertyNames
    }, {
        option: '-f, --force'
    });
}, _TenantPeopleProfileCardPropertyRemoveCommand_initValidators = function _TenantPeopleProfileCardPropertyRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (profileCardPropertyNames.every(n => n.toLowerCase() !== args.options.name.toLowerCase())) {
            return `${args.options.name} is not a valid value for name. Allowed values are ${profileCardPropertyNames.join(', ')}`;
        }
        return true;
    });
};
export default new TenantPeopleProfileCardPropertyRemoveCommand();
//# sourceMappingURL=people-profilecardproperty-remove.js.map