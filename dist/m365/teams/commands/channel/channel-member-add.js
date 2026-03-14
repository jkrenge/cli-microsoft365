var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChannelMemberAddCommand_instances, _TeamsChannelMemberAddCommand_initTelemetry, _TeamsChannelMemberAddCommand_initOptions, _TeamsChannelMemberAddCommand_initValidators, _TeamsChannelMemberAddCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class TeamsChannelMemberAddCommand extends GraphCommand {
    get name() {
        return commands.CHANNEL_MEMBER_ADD;
    }
    get description() {
        return 'Adds a specified member in the specified Microsoft Teams private or shared team channel';
    }
    constructor() {
        super();
        _TeamsChannelMemberAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberAddCommand_instances, "m", _TeamsChannelMemberAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberAddCommand_instances, "m", _TeamsChannelMemberAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberAddCommand_instances, "m", _TeamsChannelMemberAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberAddCommand_instances, "m", _TeamsChannelMemberAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const teamId = await this.getTeamId(args);
            const channelId = await this.getChannelId(teamId, args);
            const userIds = await this.getUserId(args);
            const endpoint = `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(channelId)}/members`;
            const roles = args.options.owner ? ["owner"] : [];
            const tasks = [];
            for (const userId of userIds) {
                tasks.push(this.addUser(userId, endpoint, roles));
            }
            await Promise.all(tasks);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async addUser(userId, endpoint, roles) {
        const requestOptions = {
            url: endpoint,
            headers: {
                'content-type': 'application/json;odata=nometadata',
                'accept': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                '@odata.type': '#microsoft.graph.aadUserConversationMember',
                'roles': roles,
                'user@odata.bind': `${this.resource}/v1.0/users('${userId}')`
            }
        };
        return request.post(requestOptions);
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
    async getChannelId(teamId, args) {
        if (args.options.channelId) {
            return args.options.channelId;
        }
        const channelRequestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.channelName)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(channelRequestOptions);
        const channelItem = response.value[0];
        if (!channelItem) {
            throw `The specified channel '${args.options.channelName}' does not exist in the Microsoft Teams team with ID '${teamId}'`;
        }
        if (channelItem.membershipType !== "private") {
            throw `The specified channel is not a private channel`;
        }
        return channelItem.id;
    }
    async getUserId(args) {
        if (args.options.userIds) {
            return args.options.userIds.split(',').map(u => u.trim());
        }
        const tasks = [];
        const userDisplayNames = args.options.userDisplayNames && args.options.userDisplayNames.split(',').map(u => u.trim());
        for (const userName of userDisplayNames) {
            tasks.push(this.getSingleUser(userName));
        }
        return Promise.all(tasks);
    }
    async getSingleUser(userDisplayName) {
        const userRequestOptions = {
            url: `${this.resource}/v1.0/users?$filter=displayName eq '${formatting.encodeQueryParameter(userDisplayName)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(userRequestOptions);
        const userItem = response.value[0];
        if (!userItem) {
            throw `The specified user '${userDisplayName}' does not exist`;
        }
        if (response.value.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
            const result = await cli.handleMultipleResultsFound(`Multiple users with display name '${userDisplayName}' found.`, resultAsKeyValuePair);
            return result.id;
        }
        return userItem.id;
    }
}
_TeamsChannelMemberAddCommand_instances = new WeakSet(), _TeamsChannelMemberAddCommand_initTelemetry = function _TeamsChannelMemberAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            channelId: typeof args.options.channelId !== 'undefined',
            channelName: typeof args.options.channelName !== 'undefined',
            userIds: typeof args.options.userIds !== 'undefined',
            userDisplayNames: typeof args.options.userDisplayNames !== 'undefined',
            owner: args.options.owner
        });
    });
}, _TeamsChannelMemberAddCommand_initOptions = function _TeamsChannelMemberAddCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '-c, --channelId [channelId]'
    }, {
        option: '--channelName [channelName]'
    }, {
        option: '--userIds [userIds]'
    }, {
        option: '--userDisplayNames [userDisplayNames]'
    }, {
        option: '--owner'
    });
}, _TeamsChannelMemberAddCommand_initValidators = function _TeamsChannelMemberAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
            return `${args.options.channelId} is not a valid Teams ChannelId`;
        }
        return true;
    });
}, _TeamsChannelMemberAddCommand_initOptionSets = function _TeamsChannelMemberAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] }, { options: ['channelId', 'channelName'] }, { options: ['userIds', 'userDisplayNames'] });
};
export default new TeamsChannelMemberAddCommand();
//# sourceMappingURL=channel-member-add.js.map