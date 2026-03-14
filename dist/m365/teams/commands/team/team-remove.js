var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsTeamRemoveCommand_instances, _TeamsTeamRemoveCommand_initTelemetry, _TeamsTeamRemoveCommand_initOptions, _TeamsTeamRemoveCommand_initValidators, _TeamsTeamRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsTeamRemoveCommand extends GraphCommand {
    get name() {
        return commands.TEAM_REMOVE;
    }
    get description() {
        return 'Removes the specified Microsoft Teams team';
    }
    constructor() {
        super();
        _TeamsTeamRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _TeamsTeamRemoveCommand_instances, "m", _TeamsTeamRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsTeamRemoveCommand_instances, "m", _TeamsTeamRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsTeamRemoveCommand_instances, "m", _TeamsTeamRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _TeamsTeamRemoveCommand_instances, "m", _TeamsTeamRemoveCommand_initOptionSets).call(this);
    }
    async getTeamId(args) {
        if (args.options.id) {
            return args.options.id;
        }
        const group = await entraGroup.getGroupByDisplayName(args.options.name);
        if (group.resourceProvisioningOptions.indexOf('Team') === -1) {
            throw `The specified team does not exist in the Microsoft Teams`;
        }
        return group.id;
    }
    async commandAction(logger, args) {
        const removeTeam = async () => {
            try {
                const teamId = await this.getTeamId(args);
                const requestOptions = {
                    url: `${this.resource}/v1.0/groups/${formatting.encodeQueryParameter(teamId)}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                await request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeTeam();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the team ${args.options.id ? args.options.id : args.options.name}?` });
            if (result) {
                await removeTeam();
            }
        }
    }
}
_TeamsTeamRemoveCommand_instances = new WeakSet(), _TeamsTeamRemoveCommand_initTelemetry = function _TeamsTeamRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString()
        });
    });
}, _TeamsTeamRemoveCommand_initOptions = function _TeamsTeamRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-f, --force'
    });
}, _TeamsTeamRemoveCommand_initValidators = function _TeamsTeamRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return true;
    });
}, _TeamsTeamRemoveCommand_initOptionSets = function _TeamsTeamRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new TeamsTeamRemoveCommand();
//# sourceMappingURL=team-remove.js.map