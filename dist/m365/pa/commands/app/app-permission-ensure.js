var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaAppPermissionEnsureCommand_instances, _a, _PaAppPermissionEnsureCommand_initTelemetry, _PaAppPermissionEnsureCommand_initOptions, _PaAppPermissionEnsureCommand_initValidators, _PaAppPermissionEnsureCommand_initOptionSets, _PaAppPermissionEnsureCommand_initTypes;
import Auth from '../../../../Auth.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
class PaAppPermissionEnsureCommand extends PowerAppsCommand {
    get name() {
        return commands.APP_PERMISSION_ENSURE;
    }
    get description() {
        return 'Assigns/updates permissions to a Power Apps app';
    }
    constructor() {
        super();
        _PaAppPermissionEnsureCommand_instances.add(this);
        __classPrivateFieldGet(this, _PaAppPermissionEnsureCommand_instances, "m", _PaAppPermissionEnsureCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionEnsureCommand_instances, "m", _PaAppPermissionEnsureCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionEnsureCommand_instances, "m", _PaAppPermissionEnsureCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionEnsureCommand_instances, "m", _PaAppPermissionEnsureCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionEnsureCommand_instances, "m", _PaAppPermissionEnsureCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Assigning/updating permissions for '${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName || (args.options.tenant && 'everyone')}' to the Power Apps app '${args.options.appName}'...`);
        }
        try {
            const principalId = await this.getPrincipalId(args.options);
            const requestOptions = {
                url: `${this.resource}/providers/Microsoft.PowerApps/${args.options.asAdmin ? `scopes/admin/environments/${args.options.environmentName}/` : ''}apps/${args.options.appName}/modifyPermissions?api-version=2022-11-01`,
                headers: {
                    accept: 'application/json'
                },
                data: {
                    put: [
                        {
                            properties: {
                                principal: {
                                    id: principalId,
                                    type: this.getPrincipalType(args.options)
                                },
                                NotifyShareTargetOption: args.options.sendInvitationMail ? 'Notify' : 'DoNotNotify',
                                roleName: args.options.roleName
                            }
                        }
                    ]
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getPrincipalType(options) {
        if (options.userId || options.userName) {
            return 'User';
        }
        if (options.groupId || options.groupName) {
            return 'Group';
        }
        return 'Tenant';
    }
    async getPrincipalId(options) {
        if (options.groupId) {
            return options.groupId;
        }
        if (options.userId) {
            return options.userId;
        }
        if (options.groupName) {
            const group = await entraGroup.getGroupByDisplayName(options.groupName);
            return group.id;
        }
        if (options.userName) {
            const userId = await entraUser.getUserIdByUpn(options.userName);
            return userId;
        }
        return accessToken.getTenantIdFromAccessToken(Auth.connection.accessTokens[Auth.defaultResource].accessToken);
    }
}
_a = PaAppPermissionEnsureCommand, _PaAppPermissionEnsureCommand_instances = new WeakSet(), _PaAppPermissionEnsureCommand_initTelemetry = function _PaAppPermissionEnsureCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            tenant: !!args.options.tenant,
            sendInvitationMail: !!args.options.sendInvitationMail,
            environmentName: typeof args.options.environmentName !== 'undefined',
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PaAppPermissionEnsureCommand_initOptions = function _PaAppPermissionEnsureCommand_initOptions() {
    this.options.unshift({
        option: '--appName <appName>'
    }, {
        option: '--roleName <roleName>',
        autocomplete: _a.roleNames
    }, {
        option: '--userId [userId]'
    }, {
        option: '--userName [userName]'
    }, {
        option: '--groupId [groupId]'
    }, {
        option: '--groupName [groupName]'
    }, {
        option: '--tenant'
    }, {
        option: '-e, --environmentName [environmentName]'
    }, {
        option: '--sendInvitationMail'
    }, {
        option: '--asAdmin'
    });
}, _PaAppPermissionEnsureCommand_initValidators = function _PaAppPermissionEnsureCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.appName)) {
            return `${args.options.appName} is not a valid GUID for appName.`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID for userId.`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID for groupId.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN) for userName.`;
        }
        if (_a.roleNames.indexOf(args.options.roleName) < 0) {
            return `${args.options.roleName} is not a valid roleName. Allowed values are: ${_a.roleNames.join(', ')}.`;
        }
        if (args.options.environmentName && !args.options.asAdmin) {
            return 'Specifying environmentName is only allowed when using asAdmin.';
        }
        if (args.options.asAdmin && !args.options.environmentName) {
            return 'Specifying asAdmin is only allowed when using environmentName.';
        }
        if (args.options.tenant && args.options.roleName !== 'CanView') {
            return 'Sharing with the entire tenant is only supported with CanView role.';
        }
        return true;
    });
}, _PaAppPermissionEnsureCommand_initOptionSets = function _PaAppPermissionEnsureCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName', 'tenant'] });
}, _PaAppPermissionEnsureCommand_initTypes = function _PaAppPermissionEnsureCommand_initTypes() {
    this.types.string.push('appName', 'roleName', 'userId', 'userName', 'groupId', 'groupName', 'environmentName');
};
PaAppPermissionEnsureCommand.roleNames = ['CanEdit', 'CanView'];
export default new PaAppPermissionEnsureCommand();
//# sourceMappingURL=app-permission-ensure.js.map