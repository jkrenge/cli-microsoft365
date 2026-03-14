var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTeamGetCommand_instances, _TeamsTeamGetCommand_initTelemetry, _TeamsTeamGetCommand_initOptions, _TeamsTeamGetCommand_initValidators, _TeamsTeamGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { teams } from '../../../../utils/teams.js';
class TeamsTeamGetCommand extends GraphCommand {
    get name() {
        return commands.TEAM_GET;
    }
    get description() {
        return 'Retrieve information about the specified Microsoft Team';
    }
    constructor() {
        super();
        _TeamsTeamGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTeamGetCommand_instances, "m", _TeamsTeamGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTeamGetCommand_instances, "m", _TeamsTeamGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTeamGetCommand_instances, "m", _TeamsTeamGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsTeamGetCommand_instances, "m", _TeamsTeamGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            let team;
            if (args.options.id) {
                const requestOptions = {
                    url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(args.options.id)}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                team = await request.get(requestOptions);
            }
            else {
                team = await teams.getTeamByDisplayName(args.options.name);
            }
            await logger.log(team);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsTeamGetCommand_instances = new WeakSet(), _TeamsTeamGetCommand_initTelemetry = function _TeamsTeamGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _TeamsTeamGetCommand_initOptions = function _TeamsTeamGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    });
}, _TeamsTeamGetCommand_initValidators = function _TeamsTeamGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _TeamsTeamGetCommand_initOptionSets = function _TeamsTeamGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new TeamsTeamGetCommand();
//# sourceMappingURL=team-get.js.map