var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChannelAddCommand_instances, _TeamsChannelAddCommand_initTelemetry, _TeamsChannelAddCommand_initOptions, _TeamsChannelAddCommand_initValidators, _TeamsChannelAddCommand_initOptionSets;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import { teams } from '../../../../utils/teams.js';
class TeamsChannelAddCommand extends GraphCommand {
    get name() {
        return commands.CHANNEL_ADD;
    }
    get description() {
        return 'Adds a channel to the specified Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsChannelAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsChannelAddCommand_instances, "m", _TeamsChannelAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChannelAddCommand_instances, "m", _TeamsChannelAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChannelAddCommand_instances, "m", _TeamsChannelAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChannelAddCommand_instances, "m", _TeamsChannelAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const teamId = await this.getTeamId(args);
            const res = await this.createChannel(args, teamId);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTeamId(args) {
        if (args.options.teamId) {
            return args.options.teamId;
        }
        return await teams.getTeamIdByDisplayName(args.options.teamName);
    }
    async createChannel(args, teamId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${teamId}/channels`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json;odata=nometadata'
            },
            data: {
                membershipType: args.options.type || 'standard',
                displayName: args.options.name
            },
            responseType: 'json'
        };
        if (args.options.type === 'private' || args.options.type === 'shared') {
            // Private and Shared channels must have at least 1 owner
            requestOptions.data.members = [
                {
                    '@odata.type': '#microsoft.graph.aadUserConversationMember',
                    'user@odata.bind': `${this.resource}/v1.0/users('${args.options.owner}')`,
                    roles: ['owner']
                }
            ];
        }
        return request.post(requestOptions);
    }
}
_TeamsChannelAddCommand_instances = new WeakSet(), _TeamsChannelAddCommand_initTelemetry = function _TeamsChannelAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            description: typeof args.options.description !== 'undefined',
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            type: args.options.type || 'standard',
            owner: typeof args.options.owner !== 'undefined'
        });
    });
}, _TeamsChannelAddCommand_initOptions = function _TeamsChannelAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '-n, --name <name>'
    }, {
        option: '-d, --description [description]'
    }, {
        option: '--type [type]',
        autocomplete: ['standard', 'private', 'shared']
    }, {
        option: '--owner [owner]'
    });
}, _TeamsChannelAddCommand_initValidators = function _TeamsChannelAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.type && ['standard', 'private', 'shared'].indexOf(args.options.type) === -1) {
            return `${args.options.type} is not a valid type value. Allowed values standard|private|shared.`;
        }
        if ((args.options.type === 'private' || args.options.type === 'shared') && !args.options.owner) {
            return `Specify owner when creating a ${args.options.type} channel.`;
        }
        if ((args.options.type !== 'private' && args.options.type !== 'shared') && args.options.owner) {
            return `Specify owner only when creating a private or shared channel.`;
        }
        return true;
    });
}, _TeamsChannelAddCommand_initOptionSets = function _TeamsChannelAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] });
};
export default new TeamsChannelAddCommand();
//# sourceMappingURL=channel-add.js.map