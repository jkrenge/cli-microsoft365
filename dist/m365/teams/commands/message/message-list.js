var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMessageListCommand_instances, _TeamsMessageListCommand_initTelemetry, _TeamsMessageListCommand_initOptions, _TeamsMessageListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsMessageListCommand extends GraphCommand {
    get name() {
        return commands.MESSAGE_LIST;
    }
    get description() {
        return 'Lists all messages from a channel in a Microsoft Teams team';
    }
    defaultProperties() {
        return ['id', 'summary', 'body'];
    }
    constructor() {
        super();
        _TeamsMessageListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMessageListCommand_instances, "m", _TeamsMessageListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMessageListCommand_instances, "m", _TeamsMessageListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMessageListCommand_instances, "m", _TeamsMessageListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const deltaExtension = args.options.since !== undefined ? `/delta?$filter=lastModifiedDateTime gt ${args.options.since}` : '';
        const endpoint = `${this.resource}/v1.0/teams/${args.options.teamId}/channels/${args.options.channelId}/messages${deltaExtension}`;
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
_TeamsMessageListCommand_instances = new WeakSet(), _TeamsMessageListCommand_initTelemetry = function _TeamsMessageListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            since: typeof args.options.since !== 'undefined'
        });
    });
}, _TeamsMessageListCommand_initOptions = function _TeamsMessageListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId <teamId>'
    }, {
        option: '-c, --channelId <channelId>'
    }, {
        option: '-s, --since [since]'
    });
}, _TeamsMessageListCommand_initValidators = function _TeamsMessageListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (!validation.isValidTeamsChannelId(args.options.channelId)) {
            return `${args.options.channelId} is not a valid Teams ChannelId`;
        }
        if (args.options.since && !validation.isValidISODateDashOnly(args.options.since)) {
            return `${args.options.since} is not a valid ISO Date (with dash separator)`;
        }
        if (args.options.since && !validation.isDateInRange(args.options.since, 8)) {
            return `${args.options.since} is not in the last 8 months (for delta messages)`;
        }
        return true;
    });
};
export default new TeamsMessageListCommand();
//# sourceMappingURL=message-list.js.map