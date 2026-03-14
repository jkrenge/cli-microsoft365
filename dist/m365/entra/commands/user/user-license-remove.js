var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraUserLicenseRemoveCommand_instances, _EntraUserLicenseRemoveCommand_initTelemetry, _EntraUserLicenseRemoveCommand_initOptions, _EntraUserLicenseRemoveCommand_initOptionSets, _EntraUserLicenseRemoveCommand_initValidators;
import commands from '../../commands.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import { cli } from '../../../../cli/cli.js';
import GraphCommand from '../../../base/GraphCommand.js';
class EntraUserLicenseRemoveCommand extends GraphCommand {
    get name() {
        return commands.USER_LICENSE_REMOVE;
    }
    get description() {
        return 'Removes a license from a user';
    }
    constructor() {
        super();
        _EntraUserLicenseRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraUserLicenseRemoveCommand_instances, "m", _EntraUserLicenseRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraUserLicenseRemoveCommand_instances, "m", _EntraUserLicenseRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraUserLicenseRemoveCommand_instances, "m", _EntraUserLicenseRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraUserLicenseRemoveCommand_instances, "m", _EntraUserLicenseRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Removing the licenses for the user '${args.options.userId || args.options.userName}'...`);
        }
        if (args.options.force) {
            await this.deleteUserLicenses(args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the licenses for the user '${args.options.userId || args.options.userName}'?` });
            if (result) {
                await this.deleteUserLicenses(args);
            }
        }
    }
    async deleteUserLicenses(args) {
        const removeLicenses = args.options.ids.split(',');
        const requestBody = { "addLicenses": [], "removeLicenses": removeLicenses };
        const requestOptions = {
            url: `${this.resource}/v1.0/users/${args.options.userId || args.options.userName}/assignLicense`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            data: requestBody,
            responseType: 'json'
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_EntraUserLicenseRemoveCommand_instances = new WeakSet(), _EntraUserLicenseRemoveCommand_initTelemetry = function _EntraUserLicenseRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _EntraUserLicenseRemoveCommand_initOptions = function _EntraUserLicenseRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--ids <ids>'
    }, {
        option: '-f, --force'
    });
}, _EntraUserLicenseRemoveCommand_initOptionSets = function _EntraUserLicenseRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName'] });
}, _EntraUserLicenseRemoveCommand_initValidators = function _EntraUserLicenseRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN)`;
        }
        if (args.options.ids && args.options.ids.split(',').some(e => !validation.isValidGuid(e))) {
            return `${args.options.ids} contains one or more invalid GUIDs`;
        }
        return true;
    });
};
export default new EntraUserLicenseRemoveCommand();
//# sourceMappingURL=user-license-remove.js.map