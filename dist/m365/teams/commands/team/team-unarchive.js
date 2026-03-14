var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTeamUnarchiveCommand_instances, _TeamsTeamUnarchiveCommand_initOptions, _TeamsTeamUnarchiveCommand_initValidators, _TeamsTeamUnarchiveCommand_initOptionSets;
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsTeamUnarchiveCommand extends GraphCommand {
    get name() {
        return commands.TEAM_UNARCHIVE;
    }
    get description() {
        return 'Restores an archived Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsTeamUnarchiveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTeamUnarchiveCommand_instances, "m", _TeamsTeamUnarchiveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTeamUnarchiveCommand_instances, "m", _TeamsTeamUnarchiveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsTeamUnarchiveCommand_instances, "m", _TeamsTeamUnarchiveCommand_initOptionSets).call(this);
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
        const endpoint = `${this.resource}/v1.0`;
        try {
            const teamId = await this.getTeamId(args);
            const requestOptions = {
                url: `${endpoint}/teams/${formatting.encodeQueryParameter(teamId)}/unarchive`,
                headers: {
                    'content-type': 'application/json;odata=nometadata',
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_TeamsTeamUnarchiveCommand_instances = new WeakSet(), _TeamsTeamUnarchiveCommand_initOptions = function _TeamsTeamUnarchiveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    });
}, _TeamsTeamUnarchiveCommand_initValidators = function _TeamsTeamUnarchiveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _TeamsTeamUnarchiveCommand_initOptionSets = function _TeamsTeamUnarchiveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new TeamsTeamUnarchiveCommand();
//# sourceMappingURL=team-unarchive.js.map