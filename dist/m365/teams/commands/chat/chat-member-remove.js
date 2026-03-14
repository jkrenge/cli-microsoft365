var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChatMemberRemoveCommand_instances, _TeamsChatMemberRemoveCommand_initTelemetry, _TeamsChatMemberRemoveCommand_initOptions, _TeamsChatMemberRemoveCommand_initValidators, _TeamsChatMemberRemoveCommand_initOptionSets;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
class TeamsChatMemberRemoveCommand extends GraphCommand {
    get name() {
        return commands.CHAT_MEMBER_REMOVE;
    }
    get description() {
        return 'Removes a member from a Microsoft Teams chat conversation';
    }
    constructor() {
        super();
        _TeamsChatMemberRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsChatMemberRemoveCommand_instances, "m", _TeamsChatMemberRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChatMemberRemoveCommand_instances, "m", _TeamsChatMemberRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChatMemberRemoveCommand_instances, "m", _TeamsChatMemberRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChatMemberRemoveCommand_instances, "m", _TeamsChatMemberRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeUserFromChat = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Removing member ${args.options.id || args.options.userId || args.options.userName} from chat with id ${args.options.chatId}...`);
                }
                const memberId = await this.getMemberId(args);
                const chatMemberRemoveOptions = {
                    url: `${this.resource}/v1.0/chats/${args.options.chatId}/members/${memberId}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    }
                };
                await request.delete(chatMemberRemoveOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeUserFromChat();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove member ${args.options.id || args.options.userId || args.options.userName} from chat with id ${args.options.chatId}?` });
            if (result) {
                await removeUserFromChat();
            }
        }
    }
    async getMemberId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const memberRequestUrl = `${this.resource}/v1.0/chats/${args.options.chatId}/members`;
        const members = await odata.getAllItems(memberRequestUrl);
        if (args.options.userName) {
            const matchingMember = members.find((memb) => memb.email.toLowerCase() === args.options.userName.toLowerCase());
            if (!matchingMember) {
                throw `Member with userName '${args.options.userName}' could not be found in the chat.`;
            }
            return matchingMember.id;
        }
        else {
            const matchingMember = members.find((memb) => memb.userId.toLowerCase() === args.options.userId.toLowerCase());
            if (!matchingMember) {
                throw `Member with userId '${args.options.userId}' could not be found in the chat.`;
            }
            return matchingMember.id;
        }
    }
}
_TeamsChatMemberRemoveCommand_instances = new WeakSet(), _TeamsChatMemberRemoveCommand_initTelemetry = function _TeamsChatMemberRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _TeamsChatMemberRemoveCommand_initOptions = function _TeamsChatMemberRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --chatId <chatId>'
    }, {
        option: '--id [id]'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '-f, --force'
    });
}, _TeamsChatMemberRemoveCommand_initValidators = function _TeamsChatMemberRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidTeamsChatId(args.options.chatId)) {
            return `${args.options.chatId} is not a valid Teams chatId.`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid userId.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name.`;
        }
        return true;
    });
}, _TeamsChatMemberRemoveCommand_initOptionSets = function _TeamsChatMemberRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'userId', 'userName'] });
};
export default new TeamsChatMemberRemoveCommand();
//# sourceMappingURL=chat-member-remove.js.map