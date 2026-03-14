var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChatMemberAddCommand_instances, _a, _TeamsChatMemberAddCommand_initTelemetry, _TeamsChatMemberAddCommand_initOptions, _TeamsChatMemberAddCommand_initValidators, _TeamsChatMemberAddCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsChatMemberAddCommand extends GraphCommand {
    get name() {
        return commands.CHAT_MEMBER_ADD;
    }
    get description() {
        return 'Adds a member to a Microsoft Teams chat conversation.';
    }
    constructor() {
        super();
        _TeamsChatMemberAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsChatMemberAddCommand_instances, "m", _TeamsChatMemberAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChatMemberAddCommand_instances, "m", _TeamsChatMemberAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChatMemberAddCommand_instances, "m", _TeamsChatMemberAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChatMemberAddCommand_instances, "m", _TeamsChatMemberAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Adding member ${args.options.userId || args.options.userName} to chat with id ${args.options.chatId}...`);
            }
            const chatMemberAddOptions = {
                url: `${this.resource}/v1.0/chats/${args.options.chatId}/members`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    '@odata.type': '#microsoft.graph.aadUserConversationMember',
                    'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName)}`,
                    visibleHistoryStartDateTime: args.options.withAllHistory ? '0001-01-01T00:00:00Z' : args.options.visibleHistoryStartDateTime,
                    roles: [args.options.role || 'owner']
                }
            };
            await request.post(chatMemberAddOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_a = TeamsChatMemberAddCommand, _TeamsChatMemberAddCommand_instances = new WeakSet(), _TeamsChatMemberAddCommand_initTelemetry = function _TeamsChatMemberAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            role: typeof args.options.role !== 'undefined',
            visibleHistoryStartDateTime: typeof args.options.visibleHistoryStartDateTime !== 'undefined',
            withAllHistory: !!args.options.withAllHistory
        });
    });
}, _TeamsChatMemberAddCommand_initOptions = function _TeamsChatMemberAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --chatId <chatId>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--role [role]',
        autocomplete: _a.roles
    }, {
        option: '--visibleHistoryStartDateTime [visibleHistoryStartDateTime]'
    }, {
        option: '--withAllHistory'
    });
}, _TeamsChatMemberAddCommand_initValidators = function _TeamsChatMemberAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidTeamsChatId(args.options.chatId)) {
            return `${args.options.chatId} is not a valid chatId.`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid userId.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName.`;
        }
        if (args.options.role && _a.roles.indexOf(args.options.role) < 0) {
            return `${args.options.role} is not a valid role. Allowed values are ${_a.roles.join(', ')}`;
        }
        if (args.options.visibleHistoryStartDateTime && !validation.isValidISODateTime(args.options.visibleHistoryStartDateTime)) {
            return `'${args.options.visibleHistoryStartDateTime}' is not a valid visibleHistoryStartDateTime.`;
        }
        return true;
    });
}, _TeamsChatMemberAddCommand_initOptionSets = function _TeamsChatMemberAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName'] }, {
        options: ['visibleHistoryStartDateTime', 'withAllHistory'],
        runsWhen: (args) => args.options.visibleHistoryStartDateTime || args.options.withAllHistory
    });
};
TeamsChatMemberAddCommand.roles = ['owner', 'guest'];
export default new TeamsChatMemberAddCommand();
//# sourceMappingURL=chat-member-add.js.map