var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsMessageRemoveCommand_instances, _TeamsMessageRemoveCommand_initTelemetry, _TeamsMessageRemoveCommand_initOptions, _TeamsMessageRemoveCommand_initValidators, _TeamsMessageRemoveCommand_initOptionSets, _TeamsMessageRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { teams } from '../../../../utils/teams.js';
import { validation } from '../../../../utils/validation.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';
class TeamsMessageRemoveCommand extends GraphDelegatedCommand {
    get name() {
        return commands.MESSAGE_REMOVE;
    }
    get description() {
        return 'Removes a message from a channel in a Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsMessageRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsMessageRemoveCommand_instances, "m", _TeamsMessageRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsMessageRemoveCommand_instances, "m", _TeamsMessageRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsMessageRemoveCommand_instances, "m", _TeamsMessageRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsMessageRemoveCommand_instances, "m", _TeamsMessageRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _TeamsMessageRemoveCommand_instances, "m", _TeamsMessageRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeTeamMessage = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Removing message '${args.options.id}' from channel '${args.options.channelId || args.options.channelName}' in team '${args.options.teamId || args.options.teamName}'.`);
                }
                const teamId = args.options.teamId || await teams.getTeamIdByDisplayName(args.options.teamName);
                const channelId = args.options.channelId || await teams.getChannelIdByDisplayName(teamId, args.options.channelName);
                const requestOptions = {
                    url: `${this.resource}/v1.0/teams/${teamId}/channels/${formatting.encodeQueryParameter(channelId)}/messages/${args.options.id}/softDelete`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                await request.post(requestOptions);
            }
            catch (err) {
                if (err.error?.error?.code === 'NotFound') {
                    this.handleError('The message was not found in the Teams channel.');
                }
                else {
                    this.handleRejectedODataJsonPromise(err);
                }
            }
        };
        if (args.options.force) {
            await removeTeamMessage();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove this message?` });
            if (result) {
                await removeTeamMessage();
            }
        }
    }
}
_TeamsMessageRemoveCommand_instances = new WeakSet(), _TeamsMessageRemoveCommand_initTelemetry = function _TeamsMessageRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            channelId: typeof args.options.channelId !== 'undefined',
            channelName: typeof args.options.channelName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _TeamsMessageRemoveCommand_initOptions = function _TeamsMessageRemoveCommand_initOptions() {
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
    }, {
        option: '-f, --force'
    });
}, _TeamsMessageRemoveCommand_initValidators = function _TeamsMessageRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `'${args.options.teamId}' is not a valid GUID for 'teamId'.`;
        }
        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
            return `'${args.options.channelId}' is not a valid ID for 'channelId'.`;
        }
        return true;
    });
}, _TeamsMessageRemoveCommand_initOptionSets = function _TeamsMessageRemoveCommand_initOptionSets() {
    this.optionSets.push({
        options: ['teamId', 'teamName']
    }, {
        options: ['channelId', 'channelName']
    });
}, _TeamsMessageRemoveCommand_initTypes = function _TeamsMessageRemoveCommand_initTypes() {
    this.types.string.push('teamId', 'teamName', 'channelId', 'channelName', 'id');
    this.types.boolean.push('force');
};
export default new TeamsMessageRemoveCommand();
//# sourceMappingURL=message-remove.js.map