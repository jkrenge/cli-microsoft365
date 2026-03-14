var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PlannerRosterMemberRemoveCommand_instances, _PlannerRosterMemberRemoveCommand_initTelemetry, _PlannerRosterMemberRemoveCommand_initOptions, _PlannerRosterMemberRemoveCommand_initOptionSets, _PlannerRosterMemberRemoveCommand_initValidators, _PlannerRosterMemberRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class PlannerRosterMemberRemoveCommand extends GraphCommand {
    get name() {
        return commands.ROSTER_MEMBER_REMOVE;
    }
    get description() {
        return 'Removes a member from a Microsoft Planner Roster';
    }
    constructor() {
        super();
        _PlannerRosterMemberRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberRemoveCommand_instances, "m", _PlannerRosterMemberRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberRemoveCommand_instances, "m", _PlannerRosterMemberRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberRemoveCommand_instances, "m", _PlannerRosterMemberRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberRemoveCommand_instances, "m", _PlannerRosterMemberRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PlannerRosterMemberRemoveCommand_instances, "m", _PlannerRosterMemberRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing member ${args.options.userName || args.options.userId} from the Microsoft Planner Roster`);
        }
        if (args.options.force) {
            await this.removeRosterMember(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove member '${args.options.userId || args.options.userName}'?` });
            if (result) {
                await this.removeRosterMember(args);
            }
        }
    }
    async getUserId(args) {
        if (args.options.userId) {
            return args.options.userId;
        }
        return entraUser.getUserIdByUpn(args.options.userName);
    }
    async removeRosterMember(args) {
        try {
            const rosterMembersContinue = await this.removeLastMemberConfirmation(args);
            if (rosterMembersContinue) {
                const userId = await this.getUserId(args);
                const requestOptions = {
                    url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members/${userId}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                await request.delete(requestOptions);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async removeLastMemberConfirmation(args) {
        if (!args.options.force) {
            const rosterMembers = await odata.getAllItems(`${this.resource}/beta/planner/rosters/${args.options.rosterId}/members?$select=Id`);
            if (rosterMembers.length === 1) {
                const result = await cli.promptForConfirmation({ message: `You are about to remove the last member of this Roster. When this happens, the Roster and all its contents will be deleted within 30 days. Are you sure you want to proceed?` });
                return result;
            }
        }
        return true;
    }
}
_PlannerRosterMemberRemoveCommand_instances = new WeakSet(), _PlannerRosterMemberRemoveCommand_initTelemetry = function _PlannerRosterMemberRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _PlannerRosterMemberRemoveCommand_initOptions = function _PlannerRosterMemberRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--rosterId <rosterId>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '-f, --force'
    });
}, _PlannerRosterMemberRemoveCommand_initOptionSets = function _PlannerRosterMemberRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName'] });
}, _PlannerRosterMemberRemoveCommand_initValidators = function _PlannerRosterMemberRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        return true;
    });
}, _PlannerRosterMemberRemoveCommand_initTypes = function _PlannerRosterMemberRemoveCommand_initTypes() {
    this.types.string.push('rosterId', 'userId', 'userName');
};
export default new PlannerRosterMemberRemoveCommand();
//# sourceMappingURL=roster-member-remove.js.map