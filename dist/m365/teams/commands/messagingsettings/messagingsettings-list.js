var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMessagingSettingsListCommand_instances, _TeamsMessagingSettingsListCommand_initOptions, _TeamsMessagingSettingsListCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
class TeamsMessagingSettingsListCommand extends GraphCommand {
    get name() {
        return commands.MESSAGINGSETTINGS_LIST;
    }
    get description() {
        return 'Lists messaging settings for a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsMessagingSettingsListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMessagingSettingsListCommand_instances, "m", _TeamsMessagingSettingsListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMessagingSettingsListCommand_instances, "m", _TeamsMessagingSettingsListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}?$select=messagingSettings`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res.messagingSettings);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsMessagingSettingsListCommand_instances = new WeakSet(), _TeamsMessagingSettingsListCommand_initOptions = function _TeamsMessagingSettingsListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    });
}, _TeamsMessagingSettingsListCommand_initValidators = function _TeamsMessagingSettingsListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        return true;
    });
};
export default new TeamsMessagingSettingsListCommand();
//# sourceMappingURL=messagingsettings-list.js.map