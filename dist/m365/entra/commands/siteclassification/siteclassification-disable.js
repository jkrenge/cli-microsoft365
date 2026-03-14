var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraSiteClassificationDisableCommand_instances, _EntraSiteClassificationDisableCommand_initTelemetry, _EntraSiteClassificationDisableCommand_initOptions;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraSiteClassificationDisableCommand extends GraphCommand {
    get name() {
        return commands.SITECLASSIFICATION_DISABLE;
    }
    get description() {
        return 'Disables site classification';
    }
    constructor() {
        super();
        _EntraSiteClassificationDisableCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraSiteClassificationDisableCommand_instances, "m", _EntraSiteClassificationDisableCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraSiteClassificationDisableCommand_instances, "m", _EntraSiteClassificationDisableCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        const disableSiteClassification = async () => {
            try {
                let requestOptions = {
                    url: `${this.resource}/v1.0/groupSettings`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                const res = await request.get(requestOptions);
                if (res.value.length === 0) {
                    throw 'Site classification is not enabled.';
                }
                const unifiedGroupSetting = res.value.filter((directorySetting) => {
                    return directorySetting.displayName === 'Group.Unified';
                });
                if (!unifiedGroupSetting || unifiedGroupSetting.length === 0) {
                    throw 'Missing DirectorySettingTemplate for "Group.Unified"';
                }
                if (!unifiedGroupSetting[0] ||
                    !unifiedGroupSetting[0].id || unifiedGroupSetting[0].id.length === 0) {
                    throw 'Missing UnifiedGroupSettting id';
                }
                requestOptions = {
                    url: `${this.resource}/v1.0/groupSettings/` + unifiedGroupSetting[0].id,
                    headers: {
                        accept: 'application/json;odata.metadata=none',
                        'content-type': 'application/json'
                    },
                    responseType: 'json'
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await disableSiteClassification();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to disable site classification?` });
            if (result) {
                await disableSiteClassification();
            }
        }
    }
}
_EntraSiteClassificationDisableCommand_instances = new WeakSet(), _EntraSiteClassificationDisableCommand_initTelemetry = function _EntraSiteClassificationDisableCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _EntraSiteClassificationDisableCommand_initOptions = function _EntraSiteClassificationDisableCommand_initOptions() {
    this.options.unshift({
        option: '-f, --force'
    });
};
export default new EntraSiteClassificationDisableCommand();
//# sourceMappingURL=siteclassification-disable.js.map