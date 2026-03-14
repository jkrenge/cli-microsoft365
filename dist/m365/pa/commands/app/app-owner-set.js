var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaAppOwnerSetCommand_instances, _a, _PaAppOwnerSetCommand_initTelemetry, _PaAppOwnerSetCommand_initOptions, _PaAppOwnerSetCommand_initValidators, _PaAppOwnerSetCommand_initOptionSets;
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
class PaAppOwnerSetCommand extends PowerAppsCommand {
    get name() {
        return commands.APP_OWNER_SET;
    }
    get description() {
        return 'Sets a new owner for a Power Apps app';
    }
    constructor() {
        super();
        _PaAppOwnerSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _PaAppOwnerSetCommand_instances, "m", _PaAppOwnerSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PaAppOwnerSetCommand_instances, "m", _PaAppOwnerSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaAppOwnerSetCommand_instances, "m", _PaAppOwnerSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PaAppOwnerSetCommand_instances, "m", _PaAppOwnerSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Setting new owner ${args.options.userId || args.options.userName} for Power Apps app ${args.options.appName}...`);
        }
        try {
            const userId = await this.getUserId(args.options);
            const requestOptions = {
                url: `${this.resource}/providers/Microsoft.PowerApps/scopes/admin/environments/${args.options.environmentName}/apps/${args.options.appName}/modifyAppOwner?api-version=2022-11-01`,
                headers: {
                    accept: 'application/json',
                    'Content-Type': 'application/json'
                },
                responseType: 'json',
                data: {
                    roleForOldAppOwner: args.options.roleForOldAppOwner,
                    newAppOwner: userId
                }
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getUserId(options) {
        if (options.userId) {
            return options.userId;
        }
        return entraUser.getUserIdByUpn(options.userName);
    }
}
_a = PaAppOwnerSetCommand, _PaAppOwnerSetCommand_instances = new WeakSet(), _PaAppOwnerSetCommand_initTelemetry = function _PaAppOwnerSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            roleForOldAppOwner: typeof args.options.roleForOldAppOwner !== 'undefined'
        });
    });
}, _PaAppOwnerSetCommand_initOptions = function _PaAppOwnerSetCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--appName <appName>'
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--roleForOldAppOwner [roleForOldAppOwner]',
        autocomplete: _a.roleForOldAppOwner
    });
}, _PaAppOwnerSetCommand_initValidators = function _PaAppOwnerSetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.appName)) {
            return `${args.options.appName} is not a valid GUID for appName`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID for userId`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid UPN for userName`;
        }
        if (args.options.roleForOldAppOwner && _a.roleForOldAppOwner.indexOf(args.options.roleForOldAppOwner) < 0) {
            return `${args.options.roleForOldAppOwner} is not a valid roleForOldAppOwner. Allowed values are: ${_a.roleForOldAppOwner.join(', ')}`;
        }
        return true;
    });
}, _PaAppOwnerSetCommand_initOptionSets = function _PaAppOwnerSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName'] });
};
PaAppOwnerSetCommand.roleForOldAppOwner = ['CanView', 'CanEdit'];
export default new PaAppOwnerSetCommand();
//# sourceMappingURL=app-owner-set.js.map