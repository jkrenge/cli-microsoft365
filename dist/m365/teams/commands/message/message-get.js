var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMessageGetCommand_instances, _TeamsMessageGetCommand_initOptions, _TeamsMessageGetCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
class TeamsMessageGetCommand extends GraphCommand {
    get name() {
        return commands.MESSAGE_GET;
    }
    get description() {
        return 'Retrieves a message from a channel in a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsMessageGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMessageGetCommand_instances, "m", _TeamsMessageGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMessageGetCommand_instances, "m", _TeamsMessageGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages/${args.options.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsMessageGetCommand_instances = new WeakSet(), _TeamsMessageGetCommand_initOptions = function _TeamsMessageGetCommand_initOptions() {
    this.options.unshift({
        option: '-t, --teamId <teamId>'
    }, {
        option: '-c, --channelId <channelId>'
    }, {
        option: '-i, --id <id>'
    });
}, _TeamsMessageGetCommand_initValidators = function _TeamsMessageGetCommand_initValidators() {
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
export default new TeamsMessageGetCommand();
//# sourceMappingURL=message-get.js.map