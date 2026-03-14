var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChatListCommand_instances, _TeamsChatListCommand_initTelemetry, _TeamsChatListCommand_initOptions, _TeamsChatListCommand_initValidators;
import auth from '../../../../Auth.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsChatListCommand extends GraphCommand {
    get name() {
        return commands.CHAT_LIST;
    }
    get description() {
        return 'Lists all chat conversations';
    }
    defaultProperties() {
        return ['id', 'topic', 'chatType'];
    }
    constructor() {
        super();
        _TeamsChatListCommand_instances.add(this);
        this.supportedTypes = ['oneOnOne', 'group', 'meeting'];
        __classPrivateFieldGet(this, _TeamsChatListCommand_instances, "m", _TeamsChatListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChatListCommand_instances, "m", _TeamsChatListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChatListCommand_instances, "m", _TeamsChatListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
        if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
            throw `The option 'userId' or 'userName' is required when obtaining chats using app only permissions`;
        }
        else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName)) {
            throw `The options 'userId' or 'userName' cannot be used when obtaining chats using delegated permissions`;
        }
        let requestUrl = `${this.resource}/v1.0/${!isAppOnlyAccessToken ? 'me' : `users/${args.options.userId || args.options.userName}`}/chats`;
        if (args.options.type) {
            requestUrl += `?$filter=chatType eq '${args.options.type}'`;
        }
        try {
            const items = await odata.getAllItems(requestUrl);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsChatListCommand_instances = new WeakSet(), _TeamsChatListCommand_initTelemetry = function _TeamsChatListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            type: args.options.type
        });
    });
}, _TeamsChatListCommand_initOptions = function _TeamsChatListCommand_initOptions() {
    this.options.unshift({
        option: '-t, --type [type]',
        autocomplete: this.supportedTypes
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _TeamsChatListCommand_initValidators = function _TeamsChatListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.type !== undefined && this.supportedTypes.indexOf(args.options.type) === -1) {
            return `${args.options.type} is not a valid chatType. Accepted values are ${this.supportedTypes.join(', ')}`;
        }
        if (args.options.userId && args.options.userName) {
            return `You can only specify either 'userId' or 'userName'`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        return true;
    });
};
export default new TeamsChatListCommand();
//# sourceMappingURL=chat-list.js.map