var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupSettingSetCommand_instances, _EntraGroupSettingSetCommand_initOptions, _EntraGroupSettingSetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupSettingSetCommand extends GraphCommand {
    get name() {
        return commands.GROUPSETTING_SET;
    }
    get description() {
        return 'Updates the particular group setting';
    }
    allowUnknownOptions() {
        return true;
    }
    constructor() {
        super();
        _EntraGroupSettingSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupSettingSetCommand_instances, "m", _EntraGroupSettingSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupSettingSetCommand_instances, "m", _EntraGroupSettingSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving group setting with id '${args.options.id}'...`);
        }
        try {
            let requestOptions = {
                url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const groupSetting = await request.get(requestOptions);
            requestOptions = {
                url: `${this.resource}/v1.0/groupSettings/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                data: {
                    displayName: groupSetting.displayName,
                    templateId: groupSetting.templateId,
                    values: this.getGroupSettingValues(args.options, groupSetting)
                },
                responseType: 'json'
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getGroupSettingValues(options, groupSetting) {
        const values = [];
        const excludeOptions = [
            'id',
            'debug',
            'verbose',
            'output'
        ];
        Object.keys(options).forEach(key => {
            if (excludeOptions.indexOf(key) === -1) {
                values.push({
                    name: key,
                    value: options[key]
                });
            }
        });
        groupSetting.values.forEach(v => {
            if (!values.find(e => e.name === v.name)) {
                values.push({
                    name: v.name,
                    value: v.value
                });
            }
        });
        return values;
    }
}
_EntraGroupSettingSetCommand_instances = new WeakSet(), _EntraGroupSettingSetCommand_initOptions = function _EntraGroupSettingSetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    });
}, _EntraGroupSettingSetCommand_initValidators = function _EntraGroupSettingSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
};
export default new EntraGroupSettingSetCommand();
//# sourceMappingURL=groupsetting-set.js.map