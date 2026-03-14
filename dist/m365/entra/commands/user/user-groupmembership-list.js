var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserGroupmembershipListCommand_instances, _EntraUserGroupmembershipListCommand_initTelemetry, _EntraUserGroupmembershipListCommand_initOptions, _EntraUserGroupmembershipListCommand_initValidators, _EntraUserGroupmembershipListCommand_initOptionSets;
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraUserGroupmembershipListCommand extends GraphCommand {
    get name() {
        return commands.USER_GROUPMEMBERSHIP_LIST;
    }
    get description() {
        return 'Retrieves all groups where the user is a member of';
    }
    constructor() {
        super();
        _EntraUserGroupmembershipListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserGroupmembershipListCommand_instances, "m", _EntraUserGroupmembershipListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserGroupmembershipListCommand_instances, "m", _EntraUserGroupmembershipListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserGroupmembershipListCommand_instances, "m", _EntraUserGroupmembershipListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraUserGroupmembershipListCommand_instances, "m", _EntraUserGroupmembershipListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let userId = args.options.userId;
        try {
            if (args.options.userName) {
                userId = await entraUser.getUserIdByUpn(args.options.userName);
            }
            else if (args.options.userEmail) {
                userId = await entraUser.getUserIdByEmail(args.options.userEmail);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/users/${userId}/getMemberGroups`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    securityEnabledOnly: !!args.options.securityEnabledOnly
                }
            };
            const groups = [];
            const results = await request.post(requestOptions);
            results.value.forEach(x => groups.push({ groupId: x }));
            await logger.log(groups);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraUserGroupmembershipListCommand_instances = new WeakSet(), _EntraUserGroupmembershipListCommand_initTelemetry = function _EntraUserGroupmembershipListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            userEmail: typeof args.options.userEmail !== 'undefined',
            securityEnabledOnly: !!args.options.securityEnabledOnly
        });
    });
}, _EntraUserGroupmembershipListCommand_initOptions = function _EntraUserGroupmembershipListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --userId [userId]'
    }, {
        option: '-n, --userName [userName]'
    }, {
        option: '-e, --userEmail [userEmail]'
    }, {
        option: '--securityEnabledOnly [securityEnabledOnly]'
    });
}, _EntraUserGroupmembershipListCommand_initValidators = function _EntraUserGroupmembershipListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name`;
        }
        if (args.options.userEmail && !validation.isValidUserPrincipalName(args.options.userEmail)) {
            return `${args.options.userEmail} is not a valid user email`;
        }
        return true;
    });
}, _EntraUserGroupmembershipListCommand_initOptionSets = function _EntraUserGroupmembershipListCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName', 'userEmail'] });
};
export default new EntraUserGroupmembershipListCommand();
//# sourceMappingURL=user-groupmembership-list.js.map