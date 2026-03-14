var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserSigninListCommand_instances, _EntraUserSigninListCommand_initTelemetry, _EntraUserSigninListCommand_initOptions, _EntraUserSigninListCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraUserSigninListCommand extends GraphCommand {
    get name() {
        return commands.USER_SIGNIN_LIST;
    }
    get description() {
        return 'Retrieves the Entra ID user sign-ins for the tenant';
    }
    constructor() {
        super();
        _EntraUserSigninListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserSigninListCommand_instances, "m", _EntraUserSigninListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserSigninListCommand_instances, "m", _EntraUserSigninListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserSigninListCommand_instances, "m", _EntraUserSigninListCommand_initValidators).call(this);
    }
    defaultProperties() {
        return ['id', 'userPrincipalName', 'appId', 'appDisplayName', 'createdDateTime'];
    }
    async commandAction(logger, args) {
        try {
            let endpoint = `${this.resource}/v1.0/auditLogs/signIns`;
            let filter = "";
            if (args.options.userName || args.options.userId) {
                filter = args.options.userId ?
                    `?$filter=userId eq '${formatting.encodeQueryParameter(args.options.userId)}'` :
                    `?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(args.options.userName)}'`;
            }
            if (args.options.appId || args.options.appDisplayName) {
                filter += filter ? " and " : "?$filter=";
                filter += args.options.appId ?
                    `appId eq '${formatting.encodeQueryParameter(args.options.appId)}'` :
                    `appDisplayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName)}'`;
            }
            endpoint += filter;
            const signins = await odata.getAllItems(endpoint);
            await logger.log(signins);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraUserSigninListCommand_instances = new WeakSet(), _EntraUserSigninListCommand_initTelemetry = function _EntraUserSigninListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userName: typeof args.options.userName !== 'undefined',
            userId: typeof args.options.userId !== 'undefined',
            appDisplayName: typeof args.options.appDisplayName !== 'undefined',
            appId: typeof args.options.appId !== 'undefined'
        });
    });
}, _EntraUserSigninListCommand_initOptions = function _EntraUserSigninListCommand_initOptions() {
    this.options.unshift({
        option: '-n, --userName [userName]'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--appDisplayName [appDisplayName]'
    }, {
        option: '--appId [appId]'
    });
}, _EntraUserSigninListCommand_initValidators = function _EntraUserSigninListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && args.options.userName) {
            return 'Specify either userId or userName, but not both';
        }
        if (args.options.appId && args.options.appDisplayName) {
            return 'Specify either appId or appDisplayName, but not both';
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid userName`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        return true;
    });
};
export default new EntraUserSigninListCommand();
//# sourceMappingURL=user-signin-list.js.map