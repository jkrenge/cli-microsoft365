var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserLicenseAddCommand_instances, _EntraUserLicenseAddCommand_initTelemetry, _EntraUserLicenseAddCommand_initOptions, _EntraUserLicenseAddCommand_initValidators, _EntraUserLicenseAddCommand_initOptionSets;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraUserLicenseAddCommand extends GraphCommand {
    get name() {
        return commands.USER_LICENSE_ADD;
    }
    get description() {
        return 'Assigns a license to a user';
    }
    constructor() {
        super();
        _EntraUserLicenseAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserLicenseAddCommand_instances, "m", _EntraUserLicenseAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserLicenseAddCommand_instances, "m", _EntraUserLicenseAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserLicenseAddCommand_instances, "m", _EntraUserLicenseAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraUserLicenseAddCommand_instances, "m", _EntraUserLicenseAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const addLicenses = args.options.ids.split(',').map(x => { return { "disabledPlans": [], "skuId": x }; });
        const requestBody = { "addLicenses": addLicenses, "removeLicenses": [] };
        const requestOptions = {
            url: `${this.resource}/v1.0/users/${args.options.userId || args.options.userName}/assignLicense`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            data: requestBody,
            responseType: 'json'
        };
        try {
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraUserLicenseAddCommand_instances = new WeakSet(), _EntraUserLicenseAddCommand_initTelemetry = function _EntraUserLicenseAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined'
        });
    });
}, _EntraUserLicenseAddCommand_initOptions = function _EntraUserLicenseAddCommand_initOptions() {
    this.options.unshift({
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--ids <ids>'
    });
}, _EntraUserLicenseAddCommand_initValidators = function _EntraUserLicenseAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.ids && args.options.ids.split(',').some(e => !validation.isValidGuid(e))) {
            return `${args.options.ids} contains one or more invalid GUIDs`;
        }
        return true;
    });
}, _EntraUserLicenseAddCommand_initOptionSets = function _EntraUserLicenseAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName'] });
};
export default new EntraUserLicenseAddCommand();
//# sourceMappingURL=user-license-add.js.map