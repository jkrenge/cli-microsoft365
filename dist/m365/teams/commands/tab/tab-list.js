var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTabListCommand_instances, _TeamsTabListCommand_initOptions, _TeamsTabListCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsTabListCommand extends GraphCommand {
    get name() {
        return commands.TAB_LIST;
    }
    get description() {
        return 'Lists tabs in the specified Microsoft Teams channel';
    }
    defaultProperties() {
        return ['id', 'displayName', 'teamsAppTabId'];
    }
    constructor() {
        super();
        _TeamsTabListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTabListCommand_instances, "m", _TeamsTabListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTabListCommand_instances, "m", _TeamsTabListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${formatting.encodeQueryParameter(args.options.channelId)}/tabs?$expand=teamsApp`;
        try {
            const items = await odata.getAllItems(endpoint);
            if (args.options.output !== 'json') {
                items.forEach(i => {
                    i.teamsAppTabId = i.teamsApp.id;
                });
            }
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsTabListCommand_instances = new WeakSet(), _TeamsTabListCommand_initOptions = function _TeamsTabListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    }, {
        option: '-c, --channelId <channelId>'
    });
}, _TeamsTabListCommand_initValidators = function _TeamsTabListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (!validation.isValidTeamsChannelId(args.options.channelId)) {
            return `${args.options.channelId} is not a valid Teams ChannelId`;
        }
        return true;
    });
};
export default new TeamsTabListCommand();
//# sourceMappingURL=tab-list.js.map