var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewSensitivityLabelListCommand_instances, _PurviewSensitivityLabelListCommand_initTelemetry, _PurviewSensitivityLabelListCommand_initOptions, _PurviewSensitivityLabelListCommand_initValidators;
import auth from '../../../../Auth.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
class PurviewSensitivityLabelListCommand extends GraphCommand {
    get name() {
        return commands.SENSITIVITYLABEL_LIST;
    }
    get description() {
        return 'Get a list of sensitivity labels';
    }
    constructor() {
        super();
        _PurviewSensitivityLabelListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewSensitivityLabelListCommand_instances, "m", _PurviewSensitivityLabelListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewSensitivityLabelListCommand_instances, "m", _PurviewSensitivityLabelListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewSensitivityLabelListCommand_instances, "m", _PurviewSensitivityLabelListCommand_initValidators).call(this);
    }
    defaultProperties() {
        return ['id', 'name', 'isActive'];
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
        if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
            this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
        }
        const requestUrl = args.options.userId || args.options.userName
            ? `${this.resource}/beta/users/${args.options.userId || args.options.userName}/security/informationProtection/sensitivityLabels`
            : `${this.resource}/beta/me/security/informationProtection/sensitivityLabels`;
        try {
            const items = await odata.getAllItems(requestUrl);
            await logger.log(items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PurviewSensitivityLabelListCommand_instances = new WeakSet(), _PurviewSensitivityLabelListCommand_initTelemetry = function _PurviewSensitivityLabelListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _PurviewSensitivityLabelListCommand_initOptions = function _PurviewSensitivityLabelListCommand_initOptions() {
    this.options.unshift({
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _PurviewSensitivityLabelListCommand_initValidators = function _PurviewSensitivityLabelListCommand_initValidators() {
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
export default new PurviewSensitivityLabelListCommand();
//# sourceMappingURL=sensitivitylabel-list.js.map