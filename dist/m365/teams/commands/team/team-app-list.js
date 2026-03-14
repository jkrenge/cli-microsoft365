var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTeamAppListCommand_instances, _TeamsTeamAppListCommand_initTelemetry, _TeamsTeamAppListCommand_initOptions, _TeamsTeamAppListCommand_initValidators, _TeamsTeamAppListCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import teamGetCommand from './team-get.js';
class TeamsTeamAppListCommand extends GraphCommand {
    get name() {
        return commands.TEAM_APP_LIST;
    }
    get description() {
        return 'List apps installed in the specified team';
    }
    defaultProperties() {
        return ['id', 'displayName', 'distributionMethod'];
    }
    constructor() {
        super();
        _TeamsTeamAppListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTeamAppListCommand_instances, "m", _TeamsTeamAppListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTeamAppListCommand_instances, "m", _TeamsTeamAppListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTeamAppListCommand_instances, "m", _TeamsTeamAppListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsTeamAppListCommand_instances, "m", _TeamsTeamAppListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Retrieving installed apps for team '${args.options.teamId || args.options.teamName}'`);
            }
            const teamId = await this.getTeamId(args);
            const res = await odata.getAllItems(`${this.resource}/v1.0/teams/${teamId}/installedApps?$expand=teamsApp,teamsAppDefinition`);
            if (!cli.shouldTrimOutput(args.options.output)) {
                await logger.log(res);
            }
            else {
                //converted to text friendly output
                await logger.log(res.map(i => {
                    return {
                        id: i.id,
                        displayName: i.teamsApp.displayName,
                        distributionMethod: i.teamsApp.distributionMethod
                    };
                }));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getTeamId(args) {
        if (args.options.teamId) {
            return args.options.teamId;
        }
        const teamGetOptions = {
            name: args.options.teamName,
            debug: this.debug,
            verbose: this.verbose
        };
        const commandOutput = await cli.executeCommandWithOutput(teamGetCommand, { options: { ...teamGetOptions, _: [] } });
        const team = JSON.parse(commandOutput.stdout);
        return team.id;
    }
}
_TeamsTeamAppListCommand_instances = new WeakSet(), _TeamsTeamAppListCommand_initTelemetry = function _TeamsTeamAppListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            teamId: typeof args.options.teamId !== 'undefined',
            teamName: typeof args.options.teamName !== 'undefined'
        });
    });
}, _TeamsTeamAppListCommand_initOptions = function _TeamsTeamAppListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --teamId [teamId]'
    }, {
        option: '-n, --teamName [teamName]'
    });
}, _TeamsTeamAppListCommand_initValidators = function _TeamsTeamAppListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        return true;
    });
}, _TeamsTeamAppListCommand_initOptionSets = function _TeamsTeamAppListCommand_initOptionSets() {
    this.optionSets.push({ options: ['teamId', 'teamName'] });
};
export default new TeamsTeamAppListCommand();
//# sourceMappingURL=team-app-list.js.map