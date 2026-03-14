var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupSettingAddCommand_instances, _EntraGroupSettingAddCommand_initOptions, _EntraGroupSettingAddCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupSettingAddCommand extends GraphCommand {
    get name() {
        return commands.GROUPSETTING_ADD;
    }
    get description() {
        return 'Creates a group setting';
    }
    constructor() {
        super();
        _EntraGroupSettingAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupSettingAddCommand_instances, "m", _EntraGroupSettingAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupSettingAddCommand_instances, "m", _EntraGroupSettingAddCommand_initValidators).call(this);
    }
    allowUnknownOptions() {
        return true;
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving group setting template with id '${args.options.templateId}'...`);
        }
        try {
            let requestOptions = {
                url: `${this.resource}/v1.0/groupSettingTemplates/${args.options.templateId}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const groupSettingTemplate = await request.get(requestOptions);
            requestOptions = {
                url: `${this.resource}/v1.0/groupSettings`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                data: {
                    templateId: args.options.templateId,
                    values: this.getGroupSettingValues(args.options, groupSettingTemplate)
                },
                responseType: 'json'
            };
            const res = await request.post(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getGroupSettingValues(options, groupSettingTemplate) {
        const values = [];
        const excludeOptions = [
            'templateId',
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
        groupSettingTemplate.values.forEach(v => {
            if (!values.find(e => e.name === v.name)) {
                values.push({
                    name: v.name,
                    value: v.defaultValue
                });
            }
        });
        return values;
    }
}
_EntraGroupSettingAddCommand_instances = new WeakSet(), _EntraGroupSettingAddCommand_initOptions = function _EntraGroupSettingAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --templateId <templateId>'
    });
}, _EntraGroupSettingAddCommand_initValidators = function _EntraGroupSettingAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.templateId)) {
            return `${args.options.templateId} is not a valid GUID`;
        }
        return true;
    });
};
export default new EntraGroupSettingAddCommand();
//# sourceMappingURL=groupsetting-add.js.map