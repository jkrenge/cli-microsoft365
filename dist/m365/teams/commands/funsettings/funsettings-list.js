var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsFunSettingsListCommand_instances, _TeamsFunSettingsListCommand_initOptions, _TeamsFunSettingsListCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsFunSettingsListCommand extends GraphCommand {
    get name() {
        return commands.FUNSETTINGS_LIST;
    }
    get description() {
        return 'Lists fun settings for the specified Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsFunSettingsListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsFunSettingsListCommand_instances, "m", _TeamsFunSettingsListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsFunSettingsListCommand_instances, "m", _TeamsFunSettingsListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Listing fun settings for team ${args.options.teamId}`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.teamId)}?$select=funSettings`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res.funSettings);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsFunSettingsListCommand_instances = new WeakSet(), _TeamsFunSettingsListCommand_initOptions = function _TeamsFunSettingsListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    });
}, _TeamsFunSettingsListCommand_initValidators = function _TeamsFunSettingsListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        return true;
    });
};
export default new TeamsFunSettingsListCommand();
//# sourceMappingURL=funsettings-list.js.map