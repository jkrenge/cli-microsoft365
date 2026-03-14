var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMemberSettingsListCommand_instances, _TeamsMemberSettingsListCommand_initOptions, _TeamsMemberSettingsListCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
class TeamsMemberSettingsListCommand extends GraphCommand {
    get name() {
        return commands.MEMBERSETTINGS_LIST;
    }
    get description() {
        return 'Lists member settings for a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsMemberSettingsListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMemberSettingsListCommand_instances, "m", _TeamsMemberSettingsListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMemberSettingsListCommand_instances, "m", _TeamsMemberSettingsListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}?$select=memberSettings`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res.memberSettings);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsMemberSettingsListCommand_instances = new WeakSet(), _TeamsMemberSettingsListCommand_initOptions = function _TeamsMemberSettingsListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    });
}, _TeamsMemberSettingsListCommand_initValidators = function _TeamsMemberSettingsListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        return true;
    });
};
export default new TeamsMemberSettingsListCommand();
//# sourceMappingURL=membersettings-list.js.map