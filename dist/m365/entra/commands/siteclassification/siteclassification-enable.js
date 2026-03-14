var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraSiteClassificationEnableCommand_instances, _EntraSiteClassificationEnableCommand_initTelemetry, _EntraSiteClassificationEnableCommand_initOptions;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraSiteClassificationEnableCommand extends GraphCommand {
    get name() {
        return commands.SITECLASSIFICATION_ENABLE;
    }
    get description() {
        return 'Enables site classification configuration';
    }
    constructor() {
        super();
        _EntraSiteClassificationEnableCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraSiteClassificationEnableCommand_instances, "m", _EntraSiteClassificationEnableCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraSiteClassificationEnableCommand_instances, "m", _EntraSiteClassificationEnableCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            let requestOptions = {
                url: `${this.resource}/v1.0/groupSettingTemplates`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            const unifiedGroupSetting = res.value.filter((directorySetting) => {
                return directorySetting.displayName === 'Group.Unified';
            });
            if (!unifiedGroupSetting ||
                unifiedGroupSetting.length === 0) {
                throw "Missing DirectorySettingTemplate for \"Group.Unified\"";
            }
            const updatedDirSettings = { values: [], templateId: unifiedGroupSetting[0].id };
            unifiedGroupSetting[0].values.forEach((directorySetting) => {
                switch (directorySetting.name) {
                    case "ClassificationList":
                        updatedDirSettings.values.push({
                            "name": directorySetting.name,
                            "value": args.options.classifications
                        });
                        break;
                    case "DefaultClassification":
                        updatedDirSettings.values.push({
                            "name": directorySetting.name,
                            "value": args.options.defaultClassification
                        });
                        break;
                    case "UsageGuidelinesUrl":
                        if (args.options.usageGuidelinesUrl) {
                            updatedDirSettings.values.push({
                                "name": directorySetting.name,
                                "value": args.options.usageGuidelinesUrl
                            });
                        }
                        else {
                            updatedDirSettings.values.push({
                                "name": directorySetting.name,
                                "value": directorySetting.defaultValue
                            });
                        }
                        break;
                    case "GuestUsageGuidelinesUrl":
                        if (args.options.guestUsageGuidelinesUrl) {
                            updatedDirSettings.values.push({
                                "name": directorySetting.name,
                                "value": args.options.guestUsageGuidelinesUrl
                            });
                        }
                        else {
                            updatedDirSettings.values.push({
                                "name": directorySetting.name,
                                "value": directorySetting.defaultValue
                            });
                        }
                        break;
                    default:
                        updatedDirSettings.values.push({
                            "name": directorySetting.name,
                            "value": directorySetting.defaultValue
                        });
                        break;
                }
            });
            requestOptions = {
                url: `${this.resource}/v1.0/groupSettings`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json',
                data: updatedDirSettings
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraSiteClassificationEnableCommand_instances = new WeakSet(), _EntraSiteClassificationEnableCommand_initTelemetry = function _EntraSiteClassificationEnableCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            usageGuidelinesUrl: typeof args.options.usageGuidelinesUrl !== 'undefined',
            guestUsageGuidelinesUrl: typeof args.options.guestUsageGuidelinesUrl !== 'undefined'
        });
    });
}, _EntraSiteClassificationEnableCommand_initOptions = function _EntraSiteClassificationEnableCommand_initOptions() {
    this.options.unshift({
        option: '-c, --classifications <classifications>'
    }, {
        option: '-d, --defaultClassification <defaultClassification>'
    }, {
        option: '-u, --usageGuidelinesUrl [usageGuidelinesUrl]'
    }, {
        option: '-g, --guestUsageGuidelinesUrl [guestUsageGuidelinesUrl]'
    });
};
export default new EntraSiteClassificationEnableCommand();
//# sourceMappingURL=siteclassification-enable.js.map