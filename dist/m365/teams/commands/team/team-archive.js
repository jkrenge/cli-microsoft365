var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTeamArchiveCommand_instances, _TeamsTeamArchiveCommand_initTelemetry, _TeamsTeamArchiveCommand_initOptions, _TeamsTeamArchiveCommand_initValidators, _TeamsTeamArchiveCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsTeamArchiveCommand extends GraphCommand {
    get name() {
        return commands.TEAM_ARCHIVE;
    }
    get description() {
        return 'Archives specified Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsTeamArchiveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTeamArchiveCommand_instances, "m", _TeamsTeamArchiveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTeamArchiveCommand_instances, "m", _TeamsTeamArchiveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTeamArchiveCommand_instances, "m", _TeamsTeamArchiveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsTeamArchiveCommand_instances, "m", _TeamsTeamArchiveCommand_initOptionSets).call(this);
    }
    async getTeamId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const group = await entraGroup.getGroupByDisplayName(args.options.name);
        if (group.resourceProvisioningOptions.indexOf('Team') === -1) {
            throw 'The specified team does not exist in the Microsoft Teams';
        }
        return group.id;
    }
    async commandAction(logger, args) {
        const siteReadOnlyForMembers = args.options.shouldSetSpoSiteReadOnlyForMembers === true;
        try {
            const teamId = await this.getTeamId(args);
            const requestOptions = {
                url: `${this.resource}/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/archive`,
                headers: {
                    'content-type': 'application/json;odata=nometadata',
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    shouldSetSpoSiteReadOnlyForMembers: siteReadOnlyForMembers
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsTeamArchiveCommand_instances = new WeakSet(), _TeamsTeamArchiveCommand_initTelemetry = function _TeamsTeamArchiveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            shouldSetSpoSiteReadOnlyForMembers: args.options.shouldSetSpoSiteReadOnlyForMembers === true
        });
    });
}, _TeamsTeamArchiveCommand_initOptions = function _TeamsTeamArchiveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '--shouldSetSpoSiteReadOnlyForMembers'
    });
}, _TeamsTeamArchiveCommand_initValidators = function _TeamsTeamArchiveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _TeamsTeamArchiveCommand_initOptionSets = function _TeamsTeamArchiveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new TeamsTeamArchiveCommand();
//# sourceMappingURL=team-archive.js.map