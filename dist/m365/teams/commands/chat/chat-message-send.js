var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChatMessageSendCommand_instances, _TeamsChatMessageSendCommand_initTelemetry, _TeamsChatMessageSendCommand_initOptions, _TeamsChatMessageSendCommand_initValidators, _TeamsChatMessageSendCommand_initOptionSets;
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { chatUtil } from './chatUtil.js';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
class TeamsChatMessageSendCommand extends GraphDelegatedCommand {
    get name() {
        return commands.CHAT_MESSAGE_SEND;
    }
    get description() {
        return 'Sends a chat message to a Microsoft Teams chat conversation.';
    }
    constructor() {
        super();
        _TeamsChatMessageSendCommand_instances.add(this);
        this.contentTypes = ['text', 'html'];
        __classPrivateFieldGet(this, _TeamsChatMessageSendCommand_instances, "m", _TeamsChatMessageSendCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChatMessageSendCommand_instances, "m", _TeamsChatMessageSendCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChatMessageSendCommand_instances, "m", _TeamsChatMessageSendCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChatMessageSendCommand_instances, "m", _TeamsChatMessageSendCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const chatId = await this.getChatId(logger, args);
            await this.sendChatMessage(chatId, args);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getChatId(logger, args) {
        if (args.options.chatId) {
            return args.options.chatId;
        }
        return args.options.userEmails
            ? this.ensureChatIdByUserEmails(args.options.userEmails)
            : this.getChatIdByName(args.options.chatName);
    }
    async ensureChatIdByUserEmails(userEmailsOption) {
        const userEmails = userEmailsOption.trim().toLowerCase().split(',').filter(e => e && e !== '');
        const currentUserEmail = accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken).toLowerCase();
        const existingChats = await chatUtil.findExistingChatsByParticipants([currentUserEmail, ...userEmails]);
        if (!existingChats || existingChats.length === 0) {
            const chat = await this.createConversation([currentUserEmail, ...userEmails]);
            return chat.id;
        }
        if (existingChats.length === 1) {
            return existingChats[0].id;
        }
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', existingChats);
        const result = await cli.handleMultipleResultsFound(`Multiple chat conversations with this name found.`, resultAsKeyValuePair);
        return result.id;
    }
    async getChatIdByName(chatName) {
        const existingChats = await chatUtil.findExistingGroupChatsByName(chatName);
        if (!existingChats || existingChats.length === 0) {
            throw 'No chat conversation was found with this name.';
        }
        if (existingChats.length === 1) {
            return existingChats[0].id;
        }
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', existingChats);
        const result = await cli.handleMultipleResultsFound(`Multiple chat conversations with this name found.`, resultAsKeyValuePair);
        return result.id;
    }
    // This Microsoft Graph API request throws an intermittent 404 exception, saying that it cannot find the principal.
    // The same behavior occurs when creating the conversation through the Graph Explorer.
    // It seems to happen when the userEmail casing does not match the casing of the actual UPN. 
    // When the first request throws an error, the second request does succeed. 
    // Therefore a retry-mechanism is implemented here. 
    async createConversation(memberEmails, retried = 0) {
        try {
            const jsonBody = {
                chatType: memberEmails.length > 2 ? 'group' : 'oneOnOne',
                members: memberEmails.map(email => {
                    return {
                        '@odata.type': '#microsoft.graph.aadUserConversationMember',
                        roles: ['owner'],
                        'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${email}`
                    };
                })
            };
            const requestOptions = {
                url: `${this.resource}/v1.0/chats`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: jsonBody
            };
            return await request.post(requestOptions);
        }
        catch (err) {
            if (err.message?.indexOf('404') > -1 && retried < 4) {
                return await this.createConversation(memberEmails, retried + 1);
            }
            throw err;
        }
    }
    async sendChatMessage(chatId, args) {
        const requestOptions = {
            url: `${this.resource}/v1.0/chats/${chatId}/messages`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            responseType: 'json',
            data: {
                body: {
                    contentType: args.options.contentType || 'text',
                    content: args.options.message
                }
            }
        };
        return request.post(requestOptions);
    }
}
_TeamsChatMessageSendCommand_instances = new WeakSet(), _TeamsChatMessageSendCommand_initTelemetry = function _TeamsChatMessageSendCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            chatId: typeof args.options.chatId !== 'undefined',
            userEmails: typeof args.options.userEmails !== 'undefined',
            chatName: typeof args.options.chatName !== 'undefined',
            contentType: args.options.contentType ?? 'text'
        });
    });
}, _TeamsChatMessageSendCommand_initOptions = function _TeamsChatMessageSendCommand_initOptions() {
    this.options.unshift({
        option: '--chatId [chatId]'
    }, {
        option: '-e, --userEmails [userEmails]'
    }, {
        option: '--chatName [chatName]'
    }, {
        option: '-m, --message <message>'
    }, {
        option: '--contentType [contentType]',
        autocomplete: this.contentTypes
    });
}, _TeamsChatMessageSendCommand_initValidators = function _TeamsChatMessageSendCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.chatId && !validation.isValidTeamsChatId(args.options.chatId)) {
            return `${args.options.chatId} is not a valid Teams ChatId.`;
        }
        if (args.options.userEmails) {
            const userEmails = args.options.userEmails.trim().toLowerCase().split(',').filter(e => e && e !== '');
            if (!userEmails || userEmails.length === 0 || userEmails.some(e => !validation.isValidUserPrincipalName(e))) {
                return `${args.options.userEmails} contains one or more invalid email addresses.`;
            }
        }
        if (args.options.contentType && !this.contentTypes.includes(args.options.contentType)) {
            return `'${args.options.contentType}' is not a valid value for option contentType. Allowed values are ${this.contentTypes.join(', ')}.`;
        }
        return true;
    });
}, _TeamsChatMessageSendCommand_initOptionSets = function _TeamsChatMessageSendCommand_initOptionSets() {
    this.optionSets.push({ options: ['chatId', 'userEmails', 'chatName'] });
};
export default new TeamsChatMessageSendCommand();
//# sourceMappingURL=chat-message-send.js.map