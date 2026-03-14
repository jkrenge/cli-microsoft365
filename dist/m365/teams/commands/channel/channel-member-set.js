var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChannelMemberSetCommand_instances, _TeamsChannelMemberSetCommand_initTelemetry, _TeamsChannelMemberSetCommand_initOptions, _TeamsChannelMemberSetCommand_initValidators, _TeamsChannelMemberSetCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class TeamsChannelMemberSetCommand extends GraphCommand {
    get name() {
        return commands.CHANNEL_MEMBER_SET;
    }
    get description() {
        return 'Updates the role of the specified member in the specified Microsoft Teams private or shared team channel';
    }
    constructor() {
        super();
        _TeamsChannelMemberSetCommand_instances.add(this);
        this.teamId = '';
        this.channelId = '';
        __classPrivateFieldGet(this, _TeamsChannelMemberSetCommand_instances, "m", _TeamsChannelMemberSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberSetCommand_instances, "m", _TeamsChannelMemberSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberSetCommand_instances, "m", _TeamsChannelMemberSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberSetCommand_instances, "m", _TeamsChannelMemberSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            this.teamId = await this.getTeamId(args);
            this.channelId = await this.getChannelId(args);
            const memberId = await this.getMemberId(args);
            const requestOptions = {
                url: `${this.resource}/v1.0/teams/${this.teamId}/channels/${this.channelId}/members/${memberId}`,
                headers: {
                    'accept': 'application/json;odata.metadata=none',
                    'Prefer': 'return=representation'
                },
                responseType: 'json',
                data: {
                    '@odata.type': '#microsoft.graph.aadUserConversationMember',
                    roles: [args.options.role]
                }
            };
            const member = await request.patch(requestOptions);
            await logger.log(member);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTeamId(args) {
        if (args.options.teamId) {
            return args.options.teamId;
        }
        const group = await entraGroup.getGroupByDisplayName(args.options.teamName);
        if (group.resourceProvisioningOptions.indexOf('Team') === -1) {
            throw `The specified team does not exist in the Microsoft Teams`;
        }
        return group.id;
    }
    async getChannelId(args) {
        if (args.options.channelId) {
            return args.options.channelId;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.channelName)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const channelItem = response.value[0];
        if (!channelItem) {
            throw 'The specified channel does not exist in the Microsoft Teams team';
        }
        if (channelItem.membershipType !== "private") {
            throw 'The specified channel is not a private channel';
        }
        return channelItem.id;
    }
    async getMemberId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${this.teamId}/channels/${this.channelId}/members`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const conversationMembers = response.value.filter(x => args.options.userId && x.userId?.toLocaleLowerCase() === args.options.userId.toLocaleLowerCase() ||
            args.options.userName && x.email?.toLocaleLowerCase() === args.options.userName.toLocaleLowerCase());
        const conversationMember = conversationMembers[0];
        if (!conversationMember) {
            throw 'The specified member does not exist in the Microsoft Teams channel';
        }
        if (conversationMembers.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', conversationMembers);
            const result = await cli.handleMultipleResultsFound(`Multiple Microsoft Teams channel members with name ${args.options.userName} found.`, resultAsKeyValuePair);
            return result.id;
        }
        return conversationMember.id;
    }
}
_TeamsChannelMemberSetCommand_instances = new WeakSet(), _TeamsChannelMemberSetCommand_initTelemetry = function _TeamsChannelMemberSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            channelId: typeof args.options.channelId !== 'undefined',
            channelName: typeof args.options.channelName !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            id: typeof args.options.id !== 'undefined'
        });
    });
}, _TeamsChannelMemberSetCommand_initOptions = function _TeamsChannelMemberSetCommand_initOptions() {
    this.options.unshift({
        option: '--teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '--channelId [channelId]'
    }, {
        option: '--channelName [channelName]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--id [id]'
    }, {
        option: '-r, --role <role>',
        autocomplete: ['owner', 'member']
    });
}, _TeamsChannelMemberSetCommand_initValidators = function _TeamsChannelMemberSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
            return `${args.options.channelId} is not a valid Teams Channel ID`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (['owner', 'member'].indexOf(args.options.role) === -1) {
            return `${args.options.role} is not a valid role value. Allowed values owner|member`;
        }
        return true;
    });
}, _TeamsChannelMemberSetCommand_initOptionSets = function _TeamsChannelMemberSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] }, { options: ['channelId', 'channelName'] }, { options: ['userName', 'userId', 'id'] });
};
export default new TeamsChannelMemberSetCommand();
//# sourceMappingURL=channel-member-set.js.map