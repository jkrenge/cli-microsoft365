var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsChannelGetCommand_instances, _TeamsChannelGetCommand_initTelemetry, _TeamsChannelGetCommand_initOptions, _TeamsChannelGetCommand_initValidators, _TeamsChannelGetCommand_initOptionSets;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { teams } from '../../../../utils/teams.js';
class TeamsChannelGetCommand extends GraphCommand {
    get name() {
        return commands.CHANNEL_GET;
    }
    get description() {
        return 'Gets information about the specific Microsoft Teams team channel';
    }
    constructor() {
        super();
        _TeamsChannelGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsChannelGetCommand_instances, "m", _TeamsChannelGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsChannelGetCommand_instances, "m", _TeamsChannelGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsChannelGetCommand_instances, "m", _TeamsChannelGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsChannelGetCommand_instances, "m", _TeamsChannelGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const teamId = args.options.teamId || await teams.getTeamIdByDisplayName(args.options.teamName);
            let channel;
            if (args.options.primary || args.options.id) {
                const requestOptions = {
                    url: `${this.resource}/v1.0/teams/${teamId}/${args.options.primary ? 'primaryChannel' : `channels/${args.options.id}`}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                channel = await request.get(requestOptions);
            }
            else {
                channel = await teams.getChannelByDisplayName(teamId, args.options.name);
            }
            await logger.log(channel);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsChannelGetCommand_instances = new WeakSet(), _TeamsChannelGetCommand_initTelemetry = function _TeamsChannelGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined',
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            primary: !!args.options.primary
        });
    });
}, _TeamsChannelGetCommand_initOptions = function _TeamsChannelGetCommand_initOptions() {
    this.options.unshift({
        option: '--teamId [teamId]'
    }, {
        option: '--teamName [teamName]'
    }, {
        option: '-i, --id [id]'
    }, {
        option: '--name [name]'
    }, {
        option: '--primary'
    });
}, _TeamsChannelGetCommand_initValidators = function _TeamsChannelGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.id && !validation.isValidTeamsChannelId(args.options.id)) {
            return `${args.options.id} is not a valid Teams channel id`;
        }
        return true;
    });
}, _TeamsChannelGetCommand_initOptionSets = function _TeamsChannelGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] }, { options: ['id', 'name', 'primary'] });
};
export default new TeamsChannelGetCommand();
//# sourceMappingURL=channel-get.js.map