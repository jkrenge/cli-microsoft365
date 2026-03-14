var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupSettingGetCommand_instances, _EntraGroupSettingGetCommand_initOptions, _EntraGroupSettingGetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupSettingGetCommand extends GraphCommand {
    get name() {
        return commands.GROUPSETTING_GET;
    }
    get description() {
        return 'Gets information about the particular group setting';
    }
    constructor() {
        super();
        _EntraGroupSettingGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupSettingGetCommand_instances, "m", _EntraGroupSettingGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupSettingGetCommand_instances, "m", _EntraGroupSettingGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
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
_EntraGroupSettingGetCommand_instances = new WeakSet(), _EntraGroupSettingGetCommand_initOptions = function _EntraGroupSettingGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
}, _EntraGroupSettingGetCommand_initValidators = function _EntraGroupSettingGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new EntraGroupSettingGetCommand();
//# sourceMappingURL=groupsetting-get.js.map