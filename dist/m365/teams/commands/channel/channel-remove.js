var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChannelRemoveCommand_instances, _TeamsChannelRemoveCommand_initTelemetry, _TeamsChannelRemoveCommand_initOptions, _TeamsChannelRemoveCommand_initValidators, _TeamsChannelRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsChannelRemoveCommand extends GraphCommand {
    get name() {
        return commands.CHANNEL_REMOVE;
    }
    get description() {
        return 'Removes the specified channel in the Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsChannelRemoveCommand_instances.add(this);
        this.teamId = "";
        __classPrivateFieldGet(this, _TeamsChannelRemoveCommand_instances, "m", _TeamsChannelRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChannelRemoveCommand_instances, "m", _TeamsChannelRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChannelRemoveCommand_instances, "m", _TeamsChannelRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChannelRemoveCommand_instances, "m", _TeamsChannelRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeChannel = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Removing channel ${args.options.id || args.options.name} from team ${args.options.teamId || args.options.teamName}`);
                }
                this.teamId = await this.getTeamId(args);
                const channelId = await this.getChannelId(args);
                const requestOptionsDelete = {
                    url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels/${formatting.encodeQueryParameter(channelId)}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                await request.delete(requestOptionsDelete);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeChannel();
        }
        else {
            const channel = args.options.name ? args.options.name : args.options.id;
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the channel ${channel} from team ${args.options.teamId || args.options.teamName}?` });
            if (result) {
                await removeChannel();
            }
        }
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
    async getChannelId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const channelRequestOptions = {
            url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(this.teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.name)}'`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const res = await request.get(channelRequestOptions);
        const channelItem = res.value[0];
        if (!channelItem) {
            throw 'The specified channel does not exist in this Microsoft Teams team';
        }
        return channelItem.id;
    }
}
_TeamsChannelRemoveCommand_instances = new WeakSet(), _TeamsChannelRemoveCommand_initTelemetry = function _TeamsChannelRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _TeamsChannelRemoveCommand_initOptions = function _TeamsChannelRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-f, --force'
    });
}, _TeamsChannelRemoveCommand_initValidators = function _TeamsChannelRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidTeamsChannelId(args.options.id)) {
            return `${args.options.id} is not a valid Teams channel id`;
        }
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        return true;
    });
}, _TeamsChannelRemoveCommand_initOptionSets = function _TeamsChannelRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] }, { options: ['teamId', 'teamName'] });
};
export default new TeamsChannelRemoveCommand();
//# sourceMappingURL=channel-remove.js.map