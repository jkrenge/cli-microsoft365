var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChannelSetCommand_instances, _TeamsChannelSetCommand_initTelemetry, _TeamsChannelSetCommand_initOptions, _TeamsChannelSetCommand_initValidators, _TeamsChannelSetCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsChannelSetCommand extends GraphCommand {
    get name() {
        return commands.CHANNEL_SET;
    }
    get description() {
        return 'Updates properties of the specified channel in the given Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsChannelSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsChannelSetCommand_instances, "m", _TeamsChannelSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChannelSetCommand_instances, "m", _TeamsChannelSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChannelSetCommand_instances, "m", _TeamsChannelSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChannelSetCommand_instances, "m", _TeamsChannelSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const teamId = await this.getTeamId(args);
            const channelId = await this.getChannelId(teamId, args);
            const data = this.mapRequestBody(args.options);
            const requestOptionsPatch = {
                url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(channelId)}`,
                headers: {
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: data
            };
            await request.patch(requestOptionsPatch);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    mapRequestBody(options) {
        const requestBody = {};
        if (options.newName) {
            requestBody.displayName = options.newName;
        }
        if (options.description) {
            requestBody.description = options.description;
        }
        return requestBody;
    }
    async getTeamId(args) {
        if (args.options.teamId) {
            return args.options.teamId;
        }
        const group = await entraGroup.getGroupByDisplayName(args.options.teamName);
        if (group.resourceProvisioningOptions.indexOf('Team') === -1) {
            throw 'The specified team does not exist';
        }
        else {
            return group.id;
        }
    }
    async getChannelId(teamId, args) {
        if (args.options.id) {
            return args.options.id;
        }
        const channelRequestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.get(channelRequestOptions);
        const channelItem = res.value[0];
        if (!channelItem) {
            throw `The specified channel does not exist in this Microsoft Teams team`;
        }
        return channelItem.id;
    }
}
_TeamsChannelSetCommand_instances = new WeakSet(), _TeamsChannelSetCommand_initTelemetry = function _TeamsChannelSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            newName: typeof args.options.newName !== 'undefined',
            description: typeof args.options.description !== 'undefined'
        });
    });
}, _TeamsChannelSetCommand_initOptions = function _TeamsChannelSetCommand_initOptions() {
    this.options.unshift({
        option: '--teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--newName [newName]'
    }, {
        option: '--description [description]'
    });
}, _TeamsChannelSetCommand_initValidators = function _TeamsChannelSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.id && !validation.isValidTeamsChannelId(args.options.id)) {
            return `${args.options.id} is not a valid Teams channel id`;
        }
        if (args.options.name && args.options.name.toLowerCase() === "general") {
            return 'General channel cannot be updated';
        }
        return true;
    });
}, _TeamsChannelSetCommand_initOptionSets = function _TeamsChannelSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] }, { options: ['teamId', 'teamName'] });
};
export default new TeamsChannelSetCommand();
//# sourceMappingURL=channel-set.js.map