var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraGroupSettingTemplateGetCommand_instances, _EntraGroupSettingTemplateGetCommand_initTelemetry, _EntraGroupSettingTemplateGetCommand_initOptions, _EntraGroupSettingTemplateGetCommand_initValidators, _EntraGroupSettingTemplateGetCommand_initOptionSets;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraGroupSettingTemplateGetCommand extends GraphCommand {
    get name() {
        return commands.GROUPSETTINGTEMPLATE_GET;
    }
    get description() {
        return 'Gets information about the specified Entra group settings template';
    }
    constructor() {
        super();
        _EntraGroupSettingTemplateGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraGroupSettingTemplateGetCommand_instances, "m", _EntraGroupSettingTemplateGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraGroupSettingTemplateGetCommand_instances, "m", _EntraGroupSettingTemplateGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraGroupSettingTemplateGetCommand_instances, "m", _EntraGroupSettingTemplateGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraGroupSettingTemplateGetCommand_instances, "m", _EntraGroupSettingTemplateGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const templates = await odata.getAllItems(`${this.resource}/v1.0/groupSettingTemplates`);
            const groupSettingTemplate = templates.filter(t => args.options.id ? t.id === args.options.id : t.displayName === args.options.displayName);
            if (groupSettingTemplate && groupSettingTemplate.length > 0) {
                await logger.log(groupSettingTemplate.pop());
            }
            else {
                throw `Resource '${(args.options.id || args.options.displayName)}' does not exist.`;
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraGroupSettingTemplateGetCommand_instances = new WeakSet(), _EntraGroupSettingTemplateGetCommand_initTelemetry = function _EntraGroupSettingTemplateGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            displayName: typeof args.options.displayName !== 'undefined'
        });
    });
}, _EntraGroupSettingTemplateGetCommand_initOptions = function _EntraGroupSettingTemplateGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --displayName [displayName]'
    });
}, _EntraGroupSettingTemplateGetCommand_initValidators = function _EntraGroupSettingTemplateGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id &&
            !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _EntraGroupSettingTemplateGetCommand_initOptionSets = function _EntraGroupSettingTemplateGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'displayName'] });
};
export default new EntraGroupSettingTemplateGetCommand();
//# sourceMappingURL=groupsettingtemplate-get.js.map