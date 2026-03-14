var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PurviewSensitivityLabelGetCommand_instances, _PurviewSensitivityLabelGetCommand_initTelemetry, _PurviewSensitivityLabelGetCommand_initOptions, _PurviewSensitivityLabelGetCommand_initValidators;
import auth from '../../../../Auth.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
class PurviewSensitivityLabelGetCommand extends GraphCommand {
    get name() {
        return commands.SENSITIVITYLABEL_GET;
    }
    get description() {
        return 'Retrieve the specified sensitivity label';
    }
    constructor() {
        super();
        _PurviewSensitivityLabelGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PurviewSensitivityLabelGetCommand_instances, "m", _PurviewSensitivityLabelGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PurviewSensitivityLabelGetCommand_instances, "m", _PurviewSensitivityLabelGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PurviewSensitivityLabelGetCommand_instances, "m", _PurviewSensitivityLabelGetCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
        if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName) {
            this.handleError(`Specify at least 'userId' or 'userName' when using application permissions.`);
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving sensitivity label with id ${args.options.id}`);
        }
        const requestUrl = args.options.userId || args.options.userName
            ? `${this.resource}/beta/users/${args.options.userId || args.options.userName}/security/informationProtection/sensitivityLabels/${args.options.id}`
            : `${this.resource}/beta/me/security/informationProtection/sensitivityLabels/${args.options.id}`;
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
_PurviewSensitivityLabelGetCommand_instances = new WeakSet(), _PurviewSensitivityLabelGetCommand_initTelemetry = function _PurviewSensitivityLabelGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _PurviewSensitivityLabelGetCommand_initOptions = function _PurviewSensitivityLabelGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    });
}, _PurviewSensitivityLabelGetCommand_initValidators = function _PurviewSensitivityLabelGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID.`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN)`;
        }
        return true;
    });
};
export default new PurviewSensitivityLabelGetCommand();
//# sourceMappingURL=sensitivitylabel-get.js.map