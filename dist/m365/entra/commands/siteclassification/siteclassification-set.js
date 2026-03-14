var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraSiteClassificationSetCommand_instances, _EntraSiteClassificationSetCommand_initTelemetry, _EntraSiteClassificationSetCommand_initOptions, _EntraSiteClassificationSetCommand_initValidators;
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraSiteClassificationSetCommand extends GraphCommand {
    get name() {
        return commands.SITECLASSIFICATION_SET;
    }
    get description() {
        return 'Updates site classification configuration';
    }
    constructor() {
        super();
        _EntraSiteClassificationSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraSiteClassificationSetCommand_instances, "m", _EntraSiteClassificationSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraSiteClassificationSetCommand_instances, "m", _EntraSiteClassificationSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraSiteClassificationSetCommand_instances, "m", _EntraSiteClassificationSetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            let requestOptions = {
                url: `${this.resource}/v1.0/groupSettings`,
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
                throw "There is no previous defined site classification which can updated.";
            }
            const updatedDirSettings = { values: [] };
            unifiedGroupSetting[0].values.forEach((directorySetting) => {
                switch (directorySetting.name) {
                    case "ClassificationList":
                        if (args.options.classifications) {
                            updatedDirSettings.values.push({
                                name: directorySetting.name,
                                value: args.options.classifications
                            });
                        }
                        else {
                            updatedDirSettings.values.push({
                                name: directorySetting.name,
                                value: directorySetting.value
                            });
                        }
                        break;
                    case "DefaultClassification":
                        if (args.options.defaultClassification) {
                            updatedDirSettings.values.push({
                                name: directorySetting.name,
                                value: args.options.defaultClassification
                            });
                        }
                        else {
                            updatedDirSettings.values.push({
                                name: directorySetting.name,
                                value: directorySetting.value
                            });
                        }
                        break;
                    case "UsageGuidelinesUrl":
                        if (args.options.usageGuidelinesUrl) {
                            updatedDirSettings.values.push({
                                name: directorySetting.name,
                                value: args.options.usageGuidelinesUrl
                            });
                        }
                        else {
                            updatedDirSettings.values.push({
                                name: directorySetting.name,
                                value: directorySetting.value
                            });
                        }
                        break;
                    case "GuestUsageGuidelinesUrl":
                        if (args.options.guestUsageGuidelinesUrl) {
                            updatedDirSettings.values.push({
                                name: directorySetting.name,
                                value: args.options.guestUsageGuidelinesUrl
                            });
                        }
                        else {
                            updatedDirSettings.values.push({
                                name: directorySetting.name,
                                value: directorySetting.value
                            });
                        }
                        break;
                    default:
                        updatedDirSettings.values.push({
                            name: directorySetting.name,
                            value: directorySetting.value
                        });
                        break;
                }
            });
            requestOptions = {
                url: `${this.resource}/v1.0/groupSettings/${unifiedGroupSetting[0].id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json',
                data: updatedDirSettings
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraSiteClassificationSetCommand_instances = new WeakSet(), _EntraSiteClassificationSetCommand_initTelemetry = function _EntraSiteClassificationSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            classifications: typeof args.options.classifications !== 'undefined',
            defaultClassification: typeof args.options.defaultClassification !== 'undefined',
            usageGuidelinesUrl: typeof args.options.usageGuidelinesUrl !== 'undefined',
            guestUsageGuidelinesUrl: typeof args.options.guestUsageGuidelinesUrl !== 'undefined'
        });
    });
}, _EntraSiteClassificationSetCommand_initOptions = function _EntraSiteClassificationSetCommand_initOptions() {
    this.options.unshift({
        option: '-c, --classifications [classifications]'
    }, {
        option: '-d, --defaultClassification [defaultClassification]'
    }, {
        option: '-u, --usageGuidelinesUrl [usageGuidelinesUrl]'
    }, {
        option: '-g, --guestUsageGuidelinesUrl [guestUsageGuidelinesUrl]'
    });
}, _EntraSiteClassificationSetCommand_initValidators = function _EntraSiteClassificationSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!args.options.classifications &&
            !args.options.defaultClassification &&
            !args.options.usageGuidelinesUrl &&
            !args.options.guestUsageGuidelinesUrl) {
            return 'Specify at least one property to update';
        }
        return true;
    });
};
export default new EntraSiteClassificationSetCommand();
//# sourceMappingURL=siteclassification-set.js.map