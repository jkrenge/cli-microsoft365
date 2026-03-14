var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaAppPermissionRemoveCommand_instances, _PaAppPermissionRemoveCommand_initTelemetry, _PaAppPermissionRemoveCommand_initOptions, _PaAppPermissionRemoveCommand_initValidators, _PaAppPermissionRemoveCommand_initOptionSets, _PaAppPermissionRemoveCommand_initTypes;
import Auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
class PaAppPermissionRemoveCommand extends PowerAppsCommand {
    get name() {
        return commands.APP_PERMISSION_REMOVE;
    }
    get description() {
        return 'Removes permissions to a Power Apps app';
    }
    constructor() {
        super();
        _PaAppPermissionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _PaAppPermissionRemoveCommand_instances, "m", _PaAppPermissionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionRemoveCommand_instances, "m", _PaAppPermissionRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionRemoveCommand_instances, "m", _PaAppPermissionRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionRemoveCommand_instances, "m", _PaAppPermissionRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _PaAppPermissionRemoveCommand_instances, "m", _PaAppPermissionRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (args.options.force) {
                await this.removeAppPermission(logger, args.options);
            }
            else {
                const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the permissions of '${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName || (args.options.tenant && 'everyone')}' from the Power App '${args.options.appName}'?` });
                if (result) {
                    await this.removeAppPermission(logger, args.options);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async removeAppPermission(logger, options) {
        if (this.verbose) {
            await logger.logToStderr(`Removing permissions for '${options.userId || options.userName || options.groupId || options.groupName || (options.tenant && 'everyone')}' for the Power Apps app ${options.appName}...`);
        }
        const principalId = await this.getPrincipalId(options);
        const requestOptions = {
            url: `${this.resource}/providers/Microsoft.PowerApps/${options.asAdmin ? `scopes/admin/environments/${options.environmentName}/` : ''}apps/${options.appName}/modifyPermissions?api-version=2022-11-01`,
            headers: {
                accept: 'application/json'
            },
            data: {
                delete: [
                    {
                        id: principalId
                    }
                ]
            },
            responseType: 'json'
        };
        await request.post(requestOptions);
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
        return `tenant-${accessToken.getTenantIdFromAccessToken(Auth.connection.accessTokens[Auth.defaultResource].accessToken)}`;
    }
}
_PaAppPermissionRemoveCommand_instances = new WeakSet(), _PaAppPermissionRemoveCommand_initTelemetry = function _PaAppPermissionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            userId: typeof args.options.userId !== 'undefined',
            userName: typeof args.options.userName !== 'undefined',
            groupId: typeof args.options.groupId !== 'undefined',
            groupName: typeof args.options.groupName !== 'undefined',
            tenant: !!args.options.tenant,
            asAdmin: !!args.options.asAdmin,
            environmentName: typeof args.options.environmentName !== 'undefined',
            force: !!args.options.force
        });
    });
}, _PaAppPermissionRemoveCommand_initOptions = function _PaAppPermissionRemoveCommand_initOptions() {
    this.options.unshift({
        option: '--appName <appName>'
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
        option: '--asAdmin'
    }, {
        option: '-e, --environmentName [environmentName]'
    }, {
        option: '-f, --force'
    });
}, _PaAppPermissionRemoveCommand_initValidators = function _PaAppPermissionRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.appName)) {
            return `${args.options.appName} is not a valid GUID for appName.`;
        }
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
            return `${args.options.userId} is not a valid GUID for userId.`;
        }
        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
            return `${args.options.userName} is not a valid user principal name (UPN) for userName.`;
        }
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
            return `${args.options.groupId} is not a valid GUID for groupId.`;
        }
        if (args.options.environmentName && !args.options.asAdmin) {
            return 'Specifying environmentName is only allowed when using asAdmin';
        }
        if (args.options.asAdmin && !args.options.environmentName) {
            return 'Specifying asAdmin is only allowed when using environmentName';
        }
        return true;
    });
}, _PaAppPermissionRemoveCommand_initOptionSets = function _PaAppPermissionRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName', 'tenant'] });
}, _PaAppPermissionRemoveCommand_initTypes = function _PaAppPermissionRemoveCommand_initTypes() {
    this.types.string.push('appName', 'userId', 'userName', 'groupId', 'groupName', 'environmentName');
};
export default new PaAppPermissionRemoveCommand();
//# sourceMappingURL=app-permission-remove.js.map