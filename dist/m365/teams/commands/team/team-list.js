var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTeamListCommand_instances, _TeamsTeamListCommand_initTelemetry, _TeamsTeamListCommand_initOptions, _TeamsTeamListCommand_initValidators, _TeamsTeamListCommand_initOptionSets, _TeamsTeamListCommand_initTypes;
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
class TeamsTeamListCommand extends GraphCommand {
    get name() {
        return commands.TEAM_LIST;
    }
    get description() {
        return 'Lists Microsoft Teams in the current tenant';
    }
    defaultProperties() {
        return ['id', 'displayName', 'isArchived', 'description'];
    }
    constructor() {
        super();
        _TeamsTeamListCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTeamListCommand_instances, "m", _TeamsTeamListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTeamListCommand_instances, "m", _TeamsTeamListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTeamListCommand_instances, "m", _TeamsTeamListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsTeamListCommand_instances, "m", _TeamsTeamListCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _TeamsTeamListCommand_instances, "m", _TeamsTeamListCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            if (!args.options.joined && !args.options.associated) {
                await logger.logToStderr(`Retrieving Microsoft Teams in the tenant...`);
            }
            else {
                const user = args.options.userId || args.options.userName || 'me';
                await logger.logToStderr(`Retrieving Microsoft Teams ${args.options.joined ? 'joined by' : 'associated with'} ${user}...`);
            }
        }
        try {
            let endpoint = `${this.resource}/v1.0`;
            if (args.options.joined || args.options.associated) {
                endpoint += args.options.userId || args.options.userName ? `/users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName)}` : '/me';
                endpoint += args.options.joined ? '/joinedTeams' : '/teamwork/associatedTeams';
                endpoint += '?$select=id';
            }
            else {
                // Get all team groups within the tenant
                endpoint += `/groups?$select=id&$filter=resourceProvisioningOptions/Any(x:x eq 'Team')`;
            }
            const groupResponse = await odata.getAllItems(endpoint);
            const groupIds = groupResponse.map(g => g.id);
            if (this.verbose) {
                await logger.logToStderr(`Retrieved ${groupIds.length} Microsoft Teams, getting additional information...`);
            }
            let teams = await this.getAllTeams(groupIds);
            // Sort teams by display name
            teams = teams.sort((x, y) => x.displayName.localeCompare(y.displayName));
            await logger.log(teams);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAllTeams(groupIds) {
        const groupBatches = [];
        for (let i = 0; groupIds.length > i; i += 20) {
            groupBatches.push(groupIds.slice(i, i + 20));
        }
        const promises = groupBatches.map(g => this.getTeamsBatch(g));
        const teams = await Promise.all(promises);
        const result = teams.reduce((prev, val) => prev.concat(val), []);
        return result;
    }
    async getTeamsBatch(groupIds) {
        const requestOptions = {
            url: `${this.resource}/v1.0/$batch`,
            headers: {
                accept: 'application/json;odata.metadata=none',
                'content-type': 'application/json'
            },
            responseType: 'json',
            data: {
                requests: groupIds.map((id, index) => ({
                    id: index.toString(),
                    method: 'GET',
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    url: `/teams/${id}`
                }))
            }
        };
        const response = await request.post(requestOptions);
        // Throw error if any of the requests failed
        for (const item of response.responses) {
            if (item.status !== 200) {
                throw item.body;
            }
        }
        return response.responses.map(r => r.body);
    }
}
_TeamsTeamListCommand_instances = new WeakSet(), _TeamsTeamListCommand_initTelemetry = function _TeamsTeamListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            joined: !!args.options.joined,
            associated: !!args.options.associated,
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _TeamsTeamListCommand_initOptions = function _TeamsTeamListCommand_initOptions() {
    this.options.unshift({
        option: '-j, --joined'
    }, {
        option: '-a, --associated'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _TeamsTeamListCommand_initValidators = function _TeamsTeamListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID for userId.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userId} is not a valid UPN for userName.`;
        }
        if ((args.options.userId || args.options.userName) && !args.options.joined && !args.options.associated) {
            return 'You must specify either joined or associated when specifying userId or userName.';
        }
        return true;
    });
}, _TeamsTeamListCommand_initOptionSets = function _TeamsTeamListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['joined', 'associated'],
        runsWhen: (args) => !!args.options.joined || !!args.options.associated
    }, {
        options: ['userId', 'userName'],
        runsWhen: (args) => typeof args.options.userId !== 'undefined' || typeof args.options.userName !== 'undefined'
    });
}, _TeamsTeamListCommand_initTypes = function _TeamsTeamListCommand_initTypes() {
    this.types.string.push('userId', 'userName');
};
export default new TeamsTeamListCommand();
//# sourceMappingURL=team-list.js.map