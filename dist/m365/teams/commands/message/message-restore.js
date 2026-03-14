var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMessageRestoreCommand_instances, _TeamsMessageRestoreCommand_initTelemetry, _TeamsMessageRestoreCommand_initOptions, _TeamsMessageRestoreCommand_initValidators, _TeamsMessageRestoreCommand_initOptionSets, _TeamsMessageRestoreCommand_initTypes;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import { teams } from '../../../../utils/teams.js';
class TeamsMessageRestoreCommand extends GraphDelegatedCommand {
    get name() {
        return commands.MESSAGE_RESTORE;
    }
    get description() {
        return 'Restores a deleted message from a channel in a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsMessageRestoreCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMessageRestoreCommand_instances, "m", _TeamsMessageRestoreCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMessageRestoreCommand_instances, "m", _TeamsMessageRestoreCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMessageRestoreCommand_instances, "m", _TeamsMessageRestoreCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsMessageRestoreCommand_instances, "m", _TeamsMessageRestoreCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _TeamsMessageRestoreCommand_instances, "m", _TeamsMessageRestoreCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Restoring deleted message '${args.options.id}' from channel '${args.options.channelId || args.options.channelName}' in the Microsoft Teams team '${args.options.teamId || args.options.teamName}'.`);
            }
            const teamId = await this.getTeamId(args.options, logger);
            const channelId = await this.getChannelId(args.options, teamId, logger);
            const requestOptions = {
                url: `${this.resource}/v1.0/teams/${teamId}/channels/${channelId}/messages/${args.options.id}/undoSoftDelete`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTeamId(options, logger) {
        if (options.teamId) {
            return options.teamId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Getting the Team ID.`);
        }
        const groupId = await teams.getTeamIdByDisplayName(options.teamName);
        return groupId;
    }
    async getChannelId(options, teamId, logger) {
        if (options.channelId) {
            return options.channelId;
        }
        if (this.verbose) {
            await logger.logToStderr(`Getting the channel ID.`);
        }
        const channelId = await teams.getChannelIdByDisplayName(teamId, options.channelName);
        return channelId;
    }
}
_TeamsMessageRestoreCommand_instances = new WeakSet(), _TeamsMessageRestoreCommand_initTelemetry = function _TeamsMessageRestoreCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            channelId: typeof args.options.channelId !== 'undefined',
            channelName: typeof args.options.channelName !== 'undefined'
        });
    });
}, _TeamsMessageRestoreCommand_initOptions = function _TeamsMessageRestoreCommand_initOptions() {
    this.options.unshift({
        option: '--teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '--channelId [channelId]'
    }, {
        option: '--channelName [channelName]'
    }, {
        option: '-i, --id <id>'
    });
}, _TeamsMessageRestoreCommand_initValidators = function _TeamsMessageRestoreCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `'${args.options.teamId}' is not a valid GUID for 'teamId'.`;
        }
        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
            return `'${args.options.channelId}' is not a valid ID for 'channelId'.`;
        }
        return true;
    });
}, _TeamsMessageRestoreCommand_initOptionSets = function _TeamsMessageRestoreCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] }, { options: ['channelId', 'channelName'] });
}, _TeamsMessageRestoreCommand_initTypes = function _TeamsMessageRestoreCommand_initTypes() {
    this.types.string.push('teamId', 'teamName', 'channelId', 'channelName', 'id');
};
export default new TeamsMessageRestoreCommand();
//# sourceMappingURL=message-restore.js.map