var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChatGetCommand_instances, _TeamsChatGetCommand_initTelemetry, _TeamsChatGetCommand_initOptions, _TeamsChatGetCommand_initValidators, _TeamsChatGetCommand_initOptionSets;
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { chatUtil } from './chatUtil.js';
import { cli } from '../../../../cli/cli.js';
class TeamsChatGetCommand extends GraphCommand {
    get name() {
        return commands.CHAT_GET;
    }
    get description() {
        return 'Get a Microsoft Teams chat conversation by id, participants or chat name.';
    }
    constructor() {
        super();
        _TeamsChatGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsChatGetCommand_instances, "m", _TeamsChatGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChatGetCommand_instances, "m", _TeamsChatGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChatGetCommand_instances, "m", _TeamsChatGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChatGetCommand_instances, "m", _TeamsChatGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const chatId = await this.getChatId(args);
            const chat = await this.getChatDetailsById(chatId);
            await logger.log(chat);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getChatId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        return args.options.participants
            ? this.getChatIdByParticipants(args.options.participants)
            : this.getChatIdByName(args.options.name);
    }
    async getChatDetailsById(id) {
        const requestOptions = {
            url: `${this.resource}/v1.0/chats/${formatting.encodeQueryParameter(id)}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    async getChatIdByParticipants(participantsString) {
        const participants = participantsString.trim().toLowerCase().split(',').filter(e => e && e !== '');
        const currentUserEmail = accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[this.resource].accessToken).toLowerCase();
        const existingChats = await chatUtil.findExistingChatsByParticipants([currentUserEmail, ...participants]);
        if (!existingChats || existingChats.length === 0) {
            throw 'No chat conversation was found with these participants.';
        }
        if (existingChats.length === 1) {
            return existingChats[0].id;
        }
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', existingChats);
        const result = await cli.handleMultipleResultsFound(`Multiple chat conversations with these participants found.`, resultAsKeyValuePair);
        return result.id;
    }
    async getChatIdByName(name) {
        const existingChats = await chatUtil.findExistingGroupChatsByName(name);
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
}
_TeamsChatGetCommand_instances = new WeakSet(), _TeamsChatGetCommand_initTelemetry = function _TeamsChatGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            participants: typeof args.options.participants !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _TeamsChatGetCommand_initOptions = function _TeamsChatGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-p, --participants [participants]'
    }, {
        option: '-n, --name [name]'
    });
}, _TeamsChatGetCommand_initValidators = function _TeamsChatGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidTeamsChatId(args.options.id)) {
            return `${args.options.id} is not a valid Teams ChatId.`;
        }
        if (args.options.participants) {
            const participants = args.options.participants.trim().toLowerCase().split(',').filter(e => e && e !== '');
            if (!participants || participants.length === 0 || participants.some(e => !validation.isValidUserPrincipalName(e))) {
                return `${args.options.participants} contains one or more invalid email addresses.`;
            }
        }
        return true;
    });
}, _TeamsChatGetCommand_initOptionSets = function _TeamsChatGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'participants', 'name'] });
};
export default new TeamsChatGetCommand();
//# sourceMappingURL=chat-get.js.map