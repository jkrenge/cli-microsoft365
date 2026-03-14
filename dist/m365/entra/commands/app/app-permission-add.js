var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppPermissionAddCommand_instances, _EntraAppPermissionAddCommand_initTelemetry, _EntraAppPermissionAddCommand_initOptions, _EntraAppPermissionAddCommand_initValidators, _EntraAppPermissionAddCommand_initOptionSets;
import { odata } from "../../../../utils/odata.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import request from "../../../../request.js";
import { validation } from "../../../../utils/validation.js";
import { entraApp } from "../../../../utils/entraApp.js";
import { entraServicePrincipal } from "../../../../utils/entraServicePrincipal.js";
var ScopeType;
(function (ScopeType) {
    ScopeType["Role"] = "Role";
    ScopeType["Scope"] = "Scope";
})(ScopeType || (ScopeType = {}));
class EntraAppPermissionAddCommand extends GraphCommand {
    get name() {
        return commands.APP_PERMISSION_ADD;
    }
    get description() {
        return 'Adds the specified application and/or delegated permissions to a specified Microsoft Entra app';
    }
    constructor() {
        super();
        _EntraAppPermissionAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppPermissionAddCommand_instances, "m", _EntraAppPermissionAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionAddCommand_instances, "m", _EntraAppPermissionAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionAddCommand_instances, "m", _EntraAppPermissionAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionAddCommand_instances, "m", _EntraAppPermissionAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const entraApp = await this.getEntraApp(args.options, logger);
            const servicePrincipals = await this.getServicePrincipals();
            const appPermissions = [];
            if (args.options.delegatedPermissions) {
                const delegatedPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.delegatedPermissions, ScopeType.Scope, appPermissions, logger);
                this.addPermissionsToResourceArray(delegatedPermissions, entraApp.requiredResourceAccess);
            }
            if (args.options.applicationPermissions) {
                const applicationPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.applicationPermissions, ScopeType.Role, appPermissions, logger);
                this.addPermissionsToResourceArray(applicationPermissions, entraApp.requiredResourceAccess);
            }
            const addPermissionsRequestOptions = {
                url: `${this.resource}/v1.0/applications/${entraApp.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    requiredResourceAccess: entraApp.requiredResourceAccess
                }
            };
            await request.patch(addPermissionsRequestOptions);
            if (args.options.grantAdminConsent) {
                let appServicePrincipal = servicePrincipals.find(sp => sp.appId === entraApp.appId);
                if (!appServicePrincipal) {
                    if (this.verbose) {
                        await logger.logToStderr(`Creating service principal for app ${entraApp.appId}...`);
                    }
                    appServicePrincipal = await entraServicePrincipal.createServicePrincipal(entraApp.appId);
                }
                await this.grantAdminConsent(appServicePrincipal, appPermissions, logger);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
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
    async getServicePrincipals() {
        return await odata.getAllItems(`${this.resource}/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);
    }
    async grantAdminConsent(servicePrincipal, appPermissions, logger) {
        for await (const permission of appPermissions) {
            if (permission.scope.length > 0) {
                if (this.verbose) {
                    await logger.logToStderr(`Granting consent for delegated permission(s) with resourceId ${permission.resourceId} and scope(s) ${permission.scope.join(' ')}`);
                }
                await this.grantOAuth2Permission(servicePrincipal.id, permission.resourceId, permission.scope.join(' '));
            }
            for await (const access of permission.resourceAccess.filter(acc => acc.type === ScopeType.Role)) {
                if (this.verbose) {
                    await logger.logToStderr(`Granting consent for application permission with resourceId ${permission.resourceId} and appRoleId ${access.id}`);
                }
                await this.addRoleToServicePrincipal(servicePrincipal.id, permission.resourceId, access.id);
            }
        }
    }
    async grantOAuth2Permission(servicePrincipalId, resourceId, scope) {
        const grantAdminConsentApplicationRequestOptions = {
            url: `${this.resource}/v1.0/oauth2PermissionGrants`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                clientId: servicePrincipalId,
                consentType: 'AllPrincipals',
                principalId: null,
                resourceId: resourceId,
                scope: scope
            }
        };
        return request.post(grantAdminConsentApplicationRequestOptions);
    }
    async addRoleToServicePrincipal(servicePrincipalId, resourceId, appRoleId) {
        const requestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments`,
            headers: {
                'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                appRoleId: appRoleId,
                principalId: servicePrincipalId,
                resourceId: resourceId
            }
        };
        return request.post(requestOptions);
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
            let permission = undefined;
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
    addPermissionsToResourceArray(permissions, existingArray) {
        permissions.forEach(resolvedRequiredResource => {
            const requiredResource = existingArray.find(api => api.resourceAppId === resolvedRequiredResource.resourceAppId);
            if (requiredResource) {
                // make sure that permission does not yet exist on the app or it will be added twice
                resolvedRequiredResource.resourceAccess.forEach(resAccess => {
                    if (!requiredResource.resourceAccess.some(res => res.id === resAccess.id)) {
                        requiredResource.resourceAccess.push(resAccess);
                    }
                });
            }
            else {
                existingArray.push(resolvedRequiredResource);
            }
        });
    }
}
_EntraAppPermissionAddCommand_instances = new WeakSet(), _EntraAppPermissionAddCommand_initTelemetry = function _EntraAppPermissionAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appName: typeof args.options.appName !== 'undefined',
            appObjectId: typeof args.options.appObjectId !== 'undefined',
            applicationPermissions: typeof args.options.applicationPermissions !== 'undefined',
            delegatedPermissions: typeof args.options.delegatedPermissions !== 'undefined',
            grantAdminConsent: !!args.options.grantAdminConsent
        });
    });
}, _EntraAppPermissionAddCommand_initOptions = function _EntraAppPermissionAddCommand_initOptions() {
    this.options.unshift({ option: '-i, --appId [appId]' }, { option: '-n, --appName [appName]' }, { option: '--appObjectId [appObjectId]' }, { option: '-a, --applicationPermissions [applicationPermissions]' }, { option: '-d, --delegatedPermissions [delegatedPermissions]' }, { option: '--grantAdminConsent' });
}, _EntraAppPermissionAddCommand_initValidators = function _EntraAppPermissionAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
            return `${args.options.appObjectId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraAppPermissionAddCommand_initOptionSets = function _EntraAppPermissionAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appName', 'appObjectId'] });
    this.optionSets.push({
        options: ['applicationPermissions', 'delegatedPermissions'],
        runsWhen: (args) => args.options.delegatedPermissions === undefined && args.options.applicationPermissions === undefined
    });
};
export default new EntraAppPermissionAddCommand();
//# sourceMappingURL=app-permission-add.js.map