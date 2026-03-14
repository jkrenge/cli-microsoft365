var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaAppPermissionListCommand_instances, _PaAppPermissionListCommand_initTelemetry, _PaAppPermissionListCommand_initOptions, _PaAppPermissionListCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
class PaAppPermissionListCommand extends PowerAppsCommand {
    get name() {
        return commands.APP_PERMISSION_LIST;
    }
    get description() {
        return 'Lists all permissions of a Power Apps app';
    }
    defaultProperties() {
        return ['roleName', 'principalId', 'principalType'];
    }
    constructor() {
        super();
        _PaAppPermissionListCommand_instances.add(this);
        this.allowedRoleNames = ['Owner', 'CanEdit', 'CanView'];
        __classPrivateFieldGet(this, _PaAppPermissionListCommand_instances, "m", _PaAppPermissionListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionListCommand_instances, "m", _PaAppPermissionListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionListCommand_instances, "m", _PaAppPermissionListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving permissions for app ${args.options.appName}${args.options.roleName !== undefined ? ` with role name ${args.options.roleName}` : ''}`);
        }
        const url = `${this.resource}/providers/Microsoft.PowerApps${args.options.asAdmin ? '/scopes/admin' : ''}${args.options.environmentName ? '/environments/' + formatting.encodeQueryParameter(args.options.environmentName) : ''}/apps/${args.options.appName}/permissions?api-version=2022-11-01`;
        try {
            let permissions = await odata.getAllItems(url);
            if (args.options.roleName) {
                permissions = permissions.filter(permission => permission.properties.roleName === args.options.roleName);
            }
            if (args.options.output !== 'json') {
                permissions.forEach(permission => {
                    permission.roleName = permission.properties.roleName;
                    permission.principalId = permission.properties.principal.id;
                    permission.principalType = permission.properties.principal.type;
                });
            }
            await logger.log(permissions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PaAppPermissionListCommand_instances = new WeakSet(), _PaAppPermissionListCommand_initTelemetry = function _PaAppPermissionListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin,
            environmentName: typeof args.options.environmentName !== 'undefined',
            roleName: typeof args.options.roleName !== 'undefined'
        });
    });
}, _PaAppPermissionListCommand_initOptions = function _PaAppPermissionListCommand_initOptions() {
    this.options.unshift({
        option: '--appName <appName>'
    }, {
        option: '--asAdmin'
    }, {
        option: '-e, --environmentName [environmentName]'
    }, {
        option: '--roleName [roleName]',
        autocomplete: this.allowedRoleNames
    });
}, _PaAppPermissionListCommand_initValidators = function _PaAppPermissionListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.appName)) {
            return `${args.options.appName} is not a valid GUID for appName.`;
        }
        if (args.options.roleName && !this.allowedRoleNames.includes(args.options.roleName)) {
            return `${args.options.roleName} is not a valid roleName. Allowed values are ${this.allowedRoleNames.join(',')}`;
        }
        if (args.options.asAdmin && !args.options.environmentName) {
            return 'Specifying the environmentName is required when using asAdmin';
        }
        if (!args.options.asAdmin && args.options.environmentName) {
            return 'Specifying environmentName is only allowed when using asAdmin';
        }
        return true;
    });
};
export default new PaAppPermissionListCommand();
//# sourceMappingURL=app-permission-list.js.map