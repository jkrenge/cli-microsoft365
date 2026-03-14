var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewSensitivityLabelPolicySettingsListCommand_instances, _PurviewSensitivityLabelPolicySettingsListCommand_initTelemetry, _PurviewSensitivityLabelPolicySettingsListCommand_initOptions, _PurviewSensitivityLabelPolicySettingsListCommand_initValidators;
import auth from '../../../../Auth.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
class PurviewSensitivityLabelPolicySettingsListCommand extends GraphCommand {
    get name() {
        return commands.SENSITIVITYLABEL_POLICYSETTINGS_LIST;
    }
    get description() {
        return 'Get a list of policy settings for a sensitivity label';
    }
    constructor() {
        super();
        _PurviewSensitivityLabelPolicySettingsListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewSensitivityLabelPolicySettingsListCommand_instances, "m", _PurviewSensitivityLabelPolicySettingsListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewSensitivityLabelPolicySettingsListCommand_instances, "m", _PurviewSensitivityLabelPolicySettingsListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewSensitivityLabelPolicySettingsListCommand_instances, "m", _PurviewSensitivityLabelPolicySettingsListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
        if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
            this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
        }
        const requestUrl = args.options.userId || args.options.userName
            ? `${this.resource}/beta/users/${args.options.userId || args.options.userName}/security/informationProtection/labelPolicySettings`
            : `${this.resource}/beta/me/security/informationProtection/labelPolicySettings`;
        const requestOptions = {
            url: requestUrl,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PurviewSensitivityLabelPolicySettingsListCommand_instances = new WeakSet(), _PurviewSensitivityLabelPolicySettingsListCommand_initTelemetry = function _PurviewSensitivityLabelPolicySettingsListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _PurviewSensitivityLabelPolicySettingsListCommand_initOptions = function _PurviewSensitivityLabelPolicySettingsListCommand_initOptions() {
    this.options.unshift({
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _PurviewSensitivityLabelPolicySettingsListCommand_initValidators = function _PurviewSensitivityLabelPolicySettingsListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN)`;
        }
        return true;
    });
};
export default new PurviewSensitivityLabelPolicySettingsListCommand();
//# sourceMappingURL=sensitivitylabel-policysettings-list.js.map