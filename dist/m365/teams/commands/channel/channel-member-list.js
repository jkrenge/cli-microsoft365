var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChannelMemberListCommand_instances, _TeamsChannelMemberListCommand_initTelemetry, _TeamsChannelMemberListCommand_initOptions, _TeamsChannelMemberListCommand_initValidators, _TeamsChannelMemberListCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsChannelMemberListCommand extends GraphCommand {
    get name() {
        return commands.CHANNEL_MEMBER_LIST;
    }
    get description() {
        return 'Lists members of the specified Microsoft Teams team channel';
    }
    defaultProperties() {
        return ['id', 'roles', 'displayName', 'userId', 'email'];
    }
    constructor() {
        super();
        _TeamsChannelMemberListCommand_instances.add(this);
        this.teamId = '';
        __classPrivateFieldGet(this, _TeamsChannelMemberListCommand_instances, "m", _TeamsChannelMemberListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberListCommand_instances, "m", _TeamsChannelMemberListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberListCommand_instances, "m", _TeamsChannelMemberListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChannelMemberListCommand_instances, "m", _TeamsChannelMemberListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            this.teamId = await this.getTeamId(args);
            const channelId = await this.getChannelId(args);
            const endpoint = `${this.resource}/v1.0/teams/${this.teamId}/channels/${channelId}/members`;
            let memberships = await odata.getAllItems(endpoint);
            if (args.options.role) {
                if (args.options.role === 'member') {
                    // Members have no role value
                    memberships = memberships.filter(i => i.roles.length === 0);
                }
                else {
                    memberships = memberships.filter(i => i.roles.indexOf(args.options.role) !== -1);
                }
            }
            await logger.log(memberships);
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
            throw 'The specified team does not exist in the Microsoft Teams';
        }
        return group.id;
    }
    async getChannelId(args) {
        if (args.options.channelId) {
            return args.options.channelId;
        }
        const channelRequestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.channelName)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(channelRequestOptions);
        const channelItem = response.value[0];
        if (!channelItem) {
            throw 'The specified channel does not exist in the Microsoft Teams team';
        }
        return channelItem.id;
    }
}
_TeamsChannelMemberListCommand_instances = new WeakSet(), _TeamsChannelMemberListCommand_initTelemetry = function _TeamsChannelMemberListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            channelId: typeof args.options.channelId !== 'undefined',
            channelName: typeof args.options.channelName !== 'undefined',
            role: typeof args.options.role
        });
    });
}, _TeamsChannelMemberListCommand_initOptions = function _TeamsChannelMemberListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '-c, --channelId [channelId]'
    }, {
        option: '--channelName [channelName]'
    }, {
        option: '-r, --role [role]',
        autocomplete: ['owner', 'member', 'guest']
    });
}, _TeamsChannelMemberListCommand_initValidators = function _TeamsChannelMemberListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
            return `${args.options.channelId} is not a valid Teams Channel ID`;
        }
        if (args.options.role) {
            if (['owner', 'member', 'guest'].indexOf(args.options.role) === -1) {
                return `${args.options.role} is not a valid role value. Allowed values owner|member|guest`;
            }
        }
        return true;
    });
}, _TeamsChannelMemberListCommand_initOptionSets = function _TeamsChannelMemberListCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] }, { options: ['channelId', 'channelName'] });
};
export default new TeamsChannelMemberListCommand();
//# sourceMappingURL=channel-member-list.js.map