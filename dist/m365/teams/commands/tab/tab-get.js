var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTabGetCommand_instances, _TeamsTabGetCommand_initTelemetry, _TeamsTabGetCommand_initOptions, _TeamsTabGetCommand_initValidators, _TeamsTabGetCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsTabGetCommand extends GraphCommand {
    get name() {
        return commands.TAB_GET;
    }
    get description() {
        return 'Gets information about the specified Microsoft Teams tab';
    }
    constructor() {
        super();
        _TeamsTabGetCommand_instances.add(this);
        this.teamId = "";
        this.channelId = "";
        __classPrivateFieldGet(this, _TeamsTabGetCommand_instances, "m", _TeamsTabGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTabGetCommand_instances, "m", _TeamsTabGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTabGetCommand_instances, "m", _TeamsTabGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsTabGetCommand_instances, "m", _TeamsTabGetCommand_initOptionSets).call(this);
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
    async getTabId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const tabRequestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels/${formatting.encodeQueryParameter(this.channelId)}/tabs?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(tabRequestOptions);
        const tabItem = response.value[0];
        if (!tabItem) {
            throw 'The specified tab does not exist in the Microsoft Teams team channel';
        }
        return tabItem.id;
    }
    async commandAction(logger, args) {
        try {
            this.teamId = await this.getTeamId(args);
            this.channelId = await this.getChannelId(args);
            const tabId = await this.getTabId(args);
            const endpoint = `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels/${formatting.encodeQueryParameter(this.channelId)}/tabs/${formatting.encodeQueryParameter(tabId)}?$expand=teamsApp`;
            const requestOptions = {
                url: endpoint,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsTabGetCommand_instances = new WeakSet(), _TeamsTabGetCommand_initTelemetry = function _TeamsTabGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            channelId: typeof args.options.channelId !== 'undefined',
            channelName: typeof args.options.channelName !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _TeamsTabGetCommand_initOptions = function _TeamsTabGetCommand_initOptions() {
    this.options.unshift({
        option: '--teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '--channelId [channelId]'
    }, {
        option: '--channelName [channelName]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    });
}, _TeamsTabGetCommand_initValidators = function _TeamsTabGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
            return `${args.options.channelId} is not a valid Teams channel id`;
        }
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.tabId} is not a valid GUID`;
        }
        return true;
    });
}, _TeamsTabGetCommand_initOptionSets = function _TeamsTabGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] }, { options: ['channelId', 'channelName'] }, { options: ['id', 'name'] });
};
export default new TeamsTabGetCommand();
//# sourceMappingURL=tab-get.js.map