var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChannelMemberRemoveCommand_instances, _TeamsChannelMemberRemoveCommand_initTelemetry, _TeamsChannelMemberRemoveCommand_initOptions, _TeamsChannelMemberRemoveCommand_initValidators, _TeamsChannelMemberRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsChannelMemberRemoveCommand extends GraphCommand {
    get name() {
        return commands.CHANNEL_MEMBER_REMOVE;
    }
    get description() {
        return 'Remove the specified member from the specified Microsoft Teams private or shared team channel';
    }
    constructor() {
        super();
        _TeamsChannelMemberRemoveCommand_instances.add(this);
        this.teamId = '';
        this.channelId = '';
        __classPrivateFieldGet(this, _TeamsChannelMemberRemoveCommand_instances, "m", _TeamsChannelMemberRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberRemoveCommand_instances, "m", _TeamsChannelMemberRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberRemoveCommand_instances, "m", _TeamsChannelMemberRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberRemoveCommand_instances, "m", _TeamsChannelMemberRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeMember = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing member ${args.options.userId || args.options.id || args.options.userName} from channel ${args.options.channelId || args.options.channelName} from team ${args.options.teamId || args.options.teamName}`);
            }
            try {
                await this.removeMemberFromChannel(args);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeMember();
        }
        else {
            const userName = args.options.userName || args.options.userId || args.options.id;
            const teamName = args.options.teamName || args.options.teamId;
            const channelName = args.options.channelName || args.options.channelId;
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the member ${userName} from the channel ${channelName} in team ${teamName}?` });
            if (result) {
                await removeMember();
            }
        }
    }
    async removeMemberFromChannel(args) {
        const teamId = await this.getTeamId(args);
        this.teamId = teamId;
        const channelId = await this.getChannelId(args);
        this.channelId = channelId;
        const memberId = await this.getMemberId(args);
        const requestOptions = {
            url: `${this.resource}/v1.0/teams/${this.teamId}/channels/${this.channelId}/members/${memberId}`,
            headers: {
                'accept': 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.delete(requestOptions);
    }
    async getTeamId(args) {
        if (args.options.teamId) {
            return args.options.teamId;
        }
        const group = await entraGroup.getGroupByDisplayName(args.options.teamName);
        if (group.resourceProvisioningOptions.indexOf('Team') === -1) {
            throw 'The specified team does not exist in the Microsoft Teams';
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
_TeamsChannelMemberRemoveCommand_instances = new WeakSet(), _TeamsChannelMemberRemoveCommand_initTelemetry = function _TeamsChannelMemberRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            channelId: typeof args.options.channelId !== 'undefined',
            channelName: typeof args.options.channelName !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _TeamsChannelMemberRemoveCommand_initOptions = function _TeamsChannelMemberRemoveCommand_initOptions() {
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
        option: '-f, --force'
    });
}, _TeamsChannelMemberRemoveCommand_initValidators = function _TeamsChannelMemberRemoveCommand_initValidators() {
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
        return true;
    });
}, _TeamsChannelMemberRemoveCommand_initOptionSets = function _TeamsChannelMemberRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] }, { options: ['channelId', 'channelName'] }, { options: ['userId', 'userName', 'id'] });
};
export default new TeamsChannelMemberRemoveCommand();
//# sourceMappingURL=channel-member-remove.js.map