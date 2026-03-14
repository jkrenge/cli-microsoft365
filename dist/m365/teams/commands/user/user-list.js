var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _TeamsUserListCommand_instances, _TeamsUserListCommand_initTelemetry, _TeamsUserListCommand_initOptions, _TeamsUserListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class TeamsUserListCommand extends GraphCommand {
    get name() {
        return commands.USER_LIST;
    }
    get description() {
        return "Lists users for the specified Microsoft Teams team";
    }
    constructor() {
        super();
        _TeamsUserListCommand_instances.add(this);
        this.items = [];
        __classPrivateFieldGet(this, _TeamsUserListCommand_instances, "m", _TeamsUserListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _TeamsUserListCommand_instances, "m", _TeamsUserListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _TeamsUserListCommand_instances, "m", _TeamsUserListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            await this.getOwners(logger, args.options.teamId);
            const items = args.options.role === "Owner" ? [] : await this.getMembersAndGuests(logger, args.options.teamId);
            this.items = this.items.concat(items);
            // Filter out duplicate added values for owners (as they are returned as members as well)
            // this aligns the output with what is displayed in the Teams UI
            this.items = this.items.filter((groupUser, index, self) => index === self.findIndex((t) => (t.id === groupUser.id && t.displayName === groupUser.displayName)));
            if (args.options.role) {
                this.items = this.items.filter(i => i.userType === args.options.role);
            }
            await logger.log(this.items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getOwners(logger, groupId) {
        const endpoint = `${this.resource}/v1.0/groups/${groupId}/owners?$select=id,displayName,userPrincipalName,userType`;
        const items = await odata.getAllItems(endpoint);
        this.items = this.items.concat(items);
        // Currently there is a bug in the Microsoft Graph that returns Owners as
        // userType 'member'. We therefore update all returned user as owner
        for (const i in this.items) {
            this.items[i].userType = "Owner";
        }
    }
    async getMembersAndGuests(logger, groupId) {
        const endpoint = `${this.resource}/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,userType`;
        return odata.getAllItems(endpoint);
    }
}
_TeamsUserListCommand_instances = new WeakSet(), _TeamsUserListCommand_initTelemetry = function _TeamsUserListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            role: args.options.role
        });
    });
}, _TeamsUserListCommand_initOptions = function _TeamsUserListCommand_initOptions() {
    this.options.unshift({
        option: "-i, --teamId <teamId>"
    }, {
        option: "-r, --role [type]",
        autocomplete: ["Owner", "Member", "Guest"]
    });
}, _TeamsUserListCommand_initValidators = function _TeamsUserListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.teamId)) {
            return `${args.options.teamId} is not a valid GUID`;
        }
        if (args.options.role) {
            if (['Owner', 'Member', 'Guest'].indexOf(args.options.role) === -1) {
                return `${args.options.role} is not a valid role value. Allowed values Owner|Member|Guest`;
            }
        }
        return true;
    });
};
export default new TeamsUserListCommand();
//# sourceMappingURL=user-list.js.map