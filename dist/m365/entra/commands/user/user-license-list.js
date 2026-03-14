var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserLicenseListCommand_instances, _EntraUserLicenseListCommand_initTelemetry, _EntraUserLicenseListCommand_initOptions, _EntraUserLicenseListCommand_initValidators, _EntraUserLicenseListCommand_initOptionSets;
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import auth from '../../../../Auth.js';
class EntraUserLicenseListCommand extends GraphCommand {
    get name() {
        return commands.USER_LICENSE_LIST;
    }
    get description() {
        return 'Lists the license details for a given user';
    }
    defaultProperties() {
        return ['id', 'skuId', 'skuPartNumber'];
    }
    constructor() {
        super();
        _EntraUserLicenseListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserLicenseListCommand_instances, "m", _EntraUserLicenseListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserLicenseListCommand_instances, "m", _EntraUserLicenseListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserLicenseListCommand_instances, "m", _EntraUserLicenseListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraUserLicenseListCommand_instances, "m", _EntraUserLicenseListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
        if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
            this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving licenses from user: ${args.options.userId || args.options.userName || 'current user'}.`);
        }
        let requestUrl = `${this.resource}/v1.0/`;
        if (args.options.userId || args.options.userName) {
            requestUrl += `users/${args.options.userId || args.options.userName}`;
        }
        else {
            requestUrl += 'me';
        }
        requestUrl += '/licenseDetails';
        try {
            const items = await odata.getAllItems(requestUrl);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraUserLicenseListCommand_instances = new WeakSet(), _EntraUserLicenseListCommand_initTelemetry = function _EntraUserLicenseListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _EntraUserLicenseListCommand_initOptions = function _EntraUserLicenseListCommand_initOptions() {
    this.options.unshift({
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _EntraUserLicenseListCommand_initValidators = function _EntraUserLicenseListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN)`;
        }
        return true;
    });
}, _EntraUserLicenseListCommand_initOptionSets = function _EntraUserLicenseListCommand_initOptionSets() {
    this.optionSets.push({
        options: ['userId', 'userName'],
        runsWhen: (args) => args.options.userId || args.options.userName
    });
};
export default new EntraUserLicenseListCommand();
//# sourceMappingURL=user-license-list.js.map