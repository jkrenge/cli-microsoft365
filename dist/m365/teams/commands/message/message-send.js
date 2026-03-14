var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMessageSendCommand_instances, _TeamsMessageSendCommand_initOptions, _TeamsMessageSendCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsMessageSendCommand extends GraphCommand {
    get name() {
        return commands.MESSAGE_SEND;
    }
    get description() {
        return 'Sends a message to a channel in a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsMessageSendCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMessageSendCommand_instances, "m", _TeamsMessageSendCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMessageSendCommand_instances, "m", _TeamsMessageSendCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const requestOptions = {
                url: `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json',
                data: {
                    body: {
                        contentType: 'html',
                        content: args.options.message
                    }
                }
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsMessageSendCommand_instances = new WeakSet(), _TeamsMessageSendCommand_initOptions = function _TeamsMessageSendCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    }, {
        option: '-c, --channelId <channelId>'
    }, {
        option: '-m, --message <message>'
    });
}, _TeamsMessageSendCommand_initValidators = function _TeamsMessageSendCommand_initValidators() {
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
export default new TeamsMessageSendCommand();
//# sourceMappingURL=message-send.js.map