var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupSettingRemoveCommand_instances, _EntraGroupSettingRemoveCommand_initTelemetry, _EntraGroupSettingRemoveCommand_initOptions, _EntraGroupSettingRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupSettingRemoveCommand extends GraphCommand {
    get name() {
        return commands.GROUPSETTING_REMOVE;
    }
    get description() {
        return 'Removes the particular group setting';
    }
    constructor() {
        super();
        _EntraGroupSettingRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupSettingRemoveCommand_instances, "m", _EntraGroupSettingRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupSettingRemoveCommand_instances, "m", _EntraGroupSettingRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupSettingRemoveCommand_instances, "m", _EntraGroupSettingRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const removeGroupSetting = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing group setting: ${args.options.id}...`);
            }
            try {
                const requestOptions = {
                    url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
                    headers: {
                        'accept': 'application/json;odata.metadata=none'
                    }
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeGroupSetting();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the group setting ${args.options.id}?` });
            if (result) {
                await removeGroupSetting();
            }
        }
    }
}
_EntraGroupSettingRemoveCommand_instances = new WeakSet(), _EntraGroupSettingRemoveCommand_initTelemetry = function _EntraGroupSettingRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _EntraGroupSettingRemoveCommand_initOptions = function _EntraGroupSettingRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-f, --force'
    });
}, _EntraGroupSettingRemoveCommand_initValidators = function _EntraGroupSettingRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new EntraGroupSettingRemoveCommand();
//# sourceMappingURL=groupsetting-remove.js.map