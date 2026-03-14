var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMessageReplyListCommand_instances, _TeamsMessageReplyListCommand_initOptions, _TeamsMessageReplyListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsMessageReplyListCommand extends GraphCommand {
    get name() {
        return commands.MESSAGE_REPLY_LIST;
    }
    get description() {
        return 'Retrieves replies to a message from a channel in a Microsoft Teams team';
    }
    defaultProperties() {
        return ['id', 'body'];
    }
    constructor() {
        super();
        _TeamsMessageReplyListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMessageReplyListCommand_instances, "m", _TeamsMessageReplyListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMessageReplyListCommand_instances, "m", _TeamsMessageReplyListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages/${args.options.messageId}/replies`;
        try {
            const items = await odata.getAllItems(endpoint);
            if (args.options.output !== 'json') {
                items.forEach(i => {
                    i.body = i.body.content;
                });
            }
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsMessageReplyListCommand_instances = new WeakSet(), _TeamsMessageReplyListCommand_initOptions = function _TeamsMessageReplyListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    }, {
        option: '-c, --channelId <channelId>'
    }, {
        option: '-m, --messageId <messageId>'
    });
}, _TeamsMessageReplyListCommand_initValidators = function _TeamsMessageReplyListCommand_initValidators() {
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
export default new TeamsMessageReplyListCommand();
//# sourceMappingURL=message-reply-list.js.map