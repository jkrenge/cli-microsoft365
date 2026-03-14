var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppPermissionRemoveCommand_instances, _EntraAppPermissionRemoveCommand_initTelemetry, _EntraAppPermissionRemoveCommand_initOptions, _EntraAppPermissionRemoveCommand_initValidators, _EntraAppPermissionRemoveCommand_initOptionSets, _EntraAppPermissionRemoveCommand_initTypes;
import { odata } from "../../../../utils/odata.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import request from "../../../../request.js";
import { validation } from "../../../../utils/validation.js";
import { cli } from "../../../../cli/cli.js";
import { entraApp } from "../../../../utils/entraApp.js";
var ScopeType;
(function (ScopeType) {
    ScopeType["Role"] = "Role";
    ScopeType["Scope"] = "Scope";
})(ScopeType || (ScopeType = {}));
class EntraAppPermissionRemoveCommand extends GraphCommand {
    get name() {
        return commands.APP_PERMISSION_REMOVE;
    }
    get description() {
        return 'Removes the specified application and/or delegated permissions from a specified Microsoft Entra app';
    }
    constructor() {
        super();
        _EntraAppPermissionRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppPermissionRemoveCommand_instances, "m", _EntraAppPermissionRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionRemoveCommand_instances, "m", _EntraAppPermissionRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionRemoveCommand_instances, "m", _EntraAppPermissionRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionRemoveCommand_instances, "m", _EntraAppPermissionRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionRemoveCommand_instances, "m", _EntraAppPermissionRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeAppPermissions = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Removing permissions from application ${args.options.appId || args.options.appObjectId || args.options.appName}...`);
                }
                const entraApp = await this.getEntraApp(args.options, logger);
                const servicePrincipals = await odata.getAllItems(`${this.resource}/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);
                const appPermissions = [];
                if (args.options.delegatedPermissions) {
                    const delegatedPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.delegatedPermissions, ScopeType.Scope, appPermissions, logger);
                    this.removePermissionsFromResourceArray(delegatedPermissions, entraApp.requiredResourceAccess);
                }
                if (args.options.applicationPermissions) {
                    const applicationPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.applicationPermissions, ScopeType.Role, appPermissions, logger);
                    this.removePermissionsFromResourceArray(applicationPermissions, entraApp.requiredResourceAccess);
                }
                for (let i = 0; i < entraApp.requiredResourceAccess.length; i++) {
                    if (entraApp.requiredResourceAccess[i].resourceAccess?.length === 0) {
                        entraApp.requiredResourceAccess.splice(i, 1);
                    }
                }
                const removePermissionRequestOptions = {
                    url: `${this.resource}/v1.0/applications/${entraApp.id}`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json',
                    data: {
                        requiredResourceAccess: entraApp.requiredResourceAccess
                    }
                };
                await request.patch(removePermissionRequestOptions);
                if (args.options.revokeAdminConsent) {
                    const appServicePrincipal = servicePrincipals.find(sp => sp.appId === entraApp.appId);
                    if (appServicePrincipal) {
                        await this.revokeAdminConsent(appServicePrincipal, appPermissions, logger);
                    }
                    else {
                        if (this.debug) {
                            await logger.logToStderr(`No service principal found for the appId: ${entraApp.appId}. Skipping revoking admin consent.`);
                        }
                    }
                }
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeAppPermissions();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the permissions from the specified application ${args.options.appId || args.options.appObjectId || args.options.appName}?` });
            if (result) {
                await removeAppPermissions();
            }
        }
    }
    async getEntraApp(options, logger) {
        const { appObjectId, appId, appName } = options;
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appObjectId ? appObjectId : (appId ? appId : appName)}...`);
        }
        if (appObjectId) {
            return await entraApp.getAppRegistrationByObjectId(appObjectId, ['id', 'appId', 'requiredResourceAccess']);
        }
        else if (appId) {
            return await entraApp.getAppRegistrationByAppId(appId, ['id', 'appId', 'requiredResourceAccess']);
        }
        else {
            return await entraApp.getAppRegistrationByAppName(appName, ['id', 'appId', 'requiredResourceAccess']);
        }
    }
    async revokeAdminConsent(servicePrincipal, appPermissions, logger) {
        // Check if contains app permissions
        let appRoleAssignments;
        let oAuth2RoleAssignments;
        if (appPermissions.some(perm => perm.resourceAccess.filter(acc => acc.type === ScopeType.Role).length > 0)) {
            // Retrieve app role assignments from service application
            appRoleAssignments = await odata.getAllItems(`${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/appRoleAssignments?$select=id,appRoleId,resourceId`);
        }
        if (appPermissions.filter(perm => perm.scope.length > 0).length > 0) {
            // Retrieve app role assignments from service application
            oAuth2RoleAssignments = await odata.getAllItems(`${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/oAuth2PermissionGrants?$select=id,resourceId,scope`);
        }
        for await (const permission of appPermissions) {
            if (permission.scope.length > 0) {
                if (this.verbose) {
                    await logger.logToStderr(`Revoking consent for delegated permission(s) with resourceId ${permission.resourceId} and scope(s) ${permission.scope.join(' ')}`);
                }
                const oAuth2RoleAssignment = oAuth2RoleAssignments.find(y => y.resourceId === permission.resourceId);
                if (oAuth2RoleAssignment) {
                    const scopes = oAuth2RoleAssignment?.scope?.split(' ');
                    permission.scope.forEach(scope => {
                        scopes.splice(scopes.indexOf(scope), 1);
                    });
                    oAuth2RoleAssignment.scope = scopes.join(' ');
                    await this.revokeOAuth2Permission(oAuth2RoleAssignment);
                }
            }
            for await (const access of permission.resourceAccess.filter(acc => acc.type === ScopeType.Role)) {
                if (this.verbose) {
                    await logger.logToStderr(`Revoking consent for application permission with resourceId ${permission.resourceId} and appRoleId ${access.id}`);
                }
                const appRoleAssignmentToRemove = appRoleAssignments.find(y => y.resourceId === permission.resourceId && y.appRoleId === access.id);
                if (appRoleAssignmentToRemove) {
                    await this.revokeApplicationPermission(servicePrincipal.id, appRoleAssignmentToRemove.id);
                }
            }
        }
    }
    async revokeOAuth2Permission(oAuth2RoleAssignment) {
        const revokeRequestOptions = {
            url: `${this.resource}/v1.0/oauth2PermissionGrants/${oAuth2RoleAssignment.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: oAuth2RoleAssignment
        };
        return request.patch(revokeRequestOptions);
    }
    async revokeApplicationPermission(servicePrincipalId, id) {
        const requestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments/${id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.delete(requestOptions);
    }
    async getRequiredResourceAccessForApis(servicePrincipals, apis, scopeType, appPermissions, logger) {
        const resolvedApis = [];
        const requestedApis = apis.split(' ').map(a => a.trim());
        for await (const api of requestedApis) {
            const pos = api.lastIndexOf('/');
            const permissionName = api.substring(pos + 1);
            const servicePrincipalName = api.substring(0, pos);
            if (this.verbose) {
                await logger.logToStderr(`Resolving ${api}...`);
                await logger.logToStderr(`Permission name: ${permissionName}`);
                await logger.logToStderr(`Service principal name: ${servicePrincipalName}`);
            }
            const servicePrincipal = servicePrincipals.find(sp => (sp.servicePrincipalNames.indexOf(servicePrincipalName) > -1 ||
                sp.servicePrincipalNames.indexOf(`${servicePrincipalName}/`) > -1));
            if (!servicePrincipal) {
                throw `Service principal ${servicePrincipalName} not found`;
            }
            let permission;
            if (scopeType === ScopeType.Scope) {
                permission = servicePrincipal.oauth2PermissionScopes.find(scope => scope.value === permissionName);
            }
            else if (scopeType === ScopeType.Role) {
                permission = servicePrincipal.appRoles.find(scope => scope.value === permissionName);
            }
            if (!permission) {
                throw `Permission ${permissionName} for service principal ${servicePrincipalName} not found`;
            }
            let resolvedApi = resolvedApis.find(a => a.resourceAppId === servicePrincipal.appId);
            if (!resolvedApi) {
                resolvedApi = {
                    resourceAppId: servicePrincipal.appId,
                    resourceAccess: []
                };
                resolvedApis.push(resolvedApi);
            }
            const resourceAccessPermission = {
                id: permission.id,
                type: scopeType
            };
            resolvedApi.resourceAccess.push(resourceAccessPermission);
            this.updateAppPermissions(servicePrincipal.id, resourceAccessPermission, permission.value, appPermissions);
        }
        return resolvedApis;
    }
    updateAppPermissions(spId, resourceAccessPermission, oAuth2PermissionValue, appPermissions) {
        let existingPermission = appPermissions.find(oauth => oauth.resourceId === spId);
        if (!existingPermission) {
            existingPermission = {
                resourceId: spId,
                resourceAccess: [],
                scope: []
            };
            appPermissions.push(existingPermission);
        }
        if (resourceAccessPermission.type === ScopeType.Scope && oAuth2PermissionValue && !existingPermission.scope.find(scp => scp === oAuth2PermissionValue)) {
            existingPermission.scope.push(oAuth2PermissionValue);
        }
        if (!existingPermission.resourceAccess.find(res => res.id === resourceAccessPermission.id)) {
            existingPermission.resourceAccess.push(resourceAccessPermission);
        }
    }
    removePermissionsFromResourceArray(permissions, existingArray) {
        permissions.forEach(resolvedRequiredResource => {
            const requiredResource = existingArray?.find(api => api.resourceAppId === resolvedRequiredResource.resourceAppId);
            if (requiredResource) {
                resolvedRequiredResource.resourceAccess.forEach(resolvedResourceAccess => {
                    requiredResource.resourceAccess = requiredResource.resourceAccess.filter(ra => ra.id !== resolvedResourceAccess.id);
                });
            }
        });
    }
}
_EntraAppPermissionRemoveCommand_instances = new WeakSet(), _EntraAppPermissionRemoveCommand_initTelemetry = function _EntraAppPermissionRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appObjectId: typeof args.options.appObjectId !== 'undefined',
            appName: typeof args.options.appName !== 'undefined',
            applicationPermissions: typeof args.options.applicationPermissions !== 'undefined',
            delegatedPermissions: typeof args.options.delegatedPermissions !== 'undefined',
            revokeAdminConsent: !!args.options.revokeAdminConsent,
            force: !!args.options.force
        });
    });
}, _EntraAppPermissionRemoveCommand_initOptions = function _EntraAppPermissionRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --appId [appId]'
    }, {
        option: '--appObjectId [appObjectId]'
    }, {
        option: '-n, --appName [appName]'
    }, {
        option: '-a, --applicationPermissions [applicationPermissions]'
    }, {
        option: '-d, --delegatedPermissions [delegatedPermissions]'
    }, {
        option: '--revokeAdminConsent'
    }, {
        option: '--force'
    });
}, _EntraAppPermissionRemoveCommand_initValidators = function _EntraAppPermissionRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
            return `${args.options.appObjectId} is not a valid GUID`;
        }
        if (args.options.delegatedPermissions) {
            const invalidPermissions = validation.isValidPermission(args.options.delegatedPermissions);
            if (Array.isArray(invalidPermissions)) {
                return `Delegated permission(s) ${invalidPermissions.join(', ')} are not fully-qualified`;
            }
        }
        if (args.options.applicationPermissions) {
            const invalidPermissions = validation.isValidPermission(args.options.applicationPermissions);
            if (Array.isArray(invalidPermissions)) {
                return `Application permission(s) ${invalidPermissions.join(', ')} are not fully-qualified`;
            }
        }
        return true;
    });
}, _EntraAppPermissionRemoveCommand_initOptionSets = function _EntraAppPermissionRemoveCommand_initOptionSets() {
    this.optionSets.push({
        options: ['appId', 'appObjectId', 'appName']
    }, {
        options: ['applicationPermissions', 'delegatedPermissions'],
        runsWhen: (args) => args.options.delegatedPermissions === undefined && args.options.applicationPermissions === undefined
    });
}, _EntraAppPermissionRemoveCommand_initTypes = function _EntraAppPermissionRemoveCommand_initTypes() {
    this.types.string.push('appId', 'appObjectId', 'appName', 'applicationPermissions', 'delegatedPermissions');
    this.types.boolean.push('revokeAdminConsent');
};
export default new EntraAppPermissionRemoveCommand();
//# sourceMappingURL=app-permission-remove.js.map