var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChatMemberListCommand_instances, _TeamsChatMemberListCommand_initOptions, _TeamsChatMemberListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsChatMemberListCommand extends GraphCommand {
    get name() {
        return commands.CHAT_MEMBER_LIST;
    }
    get description() {
        return 'Lists all members from a chat';
    }
    defaultProperties() {
        return ['userId', 'displayName', 'email'];
    }
    constructor() {
        super();
        _TeamsChatMemberListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsChatMemberListCommand_instances, "m", _TeamsChatMemberListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChatMemberListCommand_instances, "m", _TeamsChatMemberListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const endpoint = `${this.resource}/v1.0/chats/${args.options.chatId}/members`;
        try {
            const items = await odata.getAllItems(endpoint);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsChatMemberListCommand_instances = new WeakSet(), _TeamsChatMemberListCommand_initOptions = function _TeamsChatMemberListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --chatId <chatId>'
    });
}, _TeamsChatMemberListCommand_initValidators = function _TeamsChatMemberListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidTeamsChatId(args.options.chatId)) {
            return `${args.options.chatId} is not a valid Teams ChatId`;
        }
        return true;
    });
};
export default new TeamsChatMemberListCommand();
//# sourceMappingURL=chat-member-list.js.map