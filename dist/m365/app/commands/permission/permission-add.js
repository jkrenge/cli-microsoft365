import { z } from 'zod';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import AppCommand, { appCommandOptions } from '../../../base/AppCommand.js';
import commands from '../../commands.js';
import { entraServicePrincipal } from '../../../../utils/entraServicePrincipal.js';
export const options = z.strictObject({
    ...appCommandOptions.shape,
    applicationPermissions: z.string().optional(),
    delegatedPermissions: z.string().optional(),
    grantAdminConsent: z.boolean().optional()
});
var ScopeType;
(function (ScopeType) {
    ScopeType["Role"] = "Role";
    ScopeType["Scope"] = "Scope";
})(ScopeType || (ScopeType = {}));
class AppPermissionAddCommand extends AppCommand {
    get name() {
        return commands.PERMISSION_ADD;
    }
    get description() {
        return 'Adds the specified application and/or delegated permissions to the current Microsoft Entra app API permissions';
    }
    get schema() {
        return options;
    }
    getRefinedSchema(schema) {
        return schema
            .refine(options => options.applicationPermissions || options.delegatedPermissions, {
            error: 'Specify at least one of applicationPermissions or delegatedPermissions, or both.',
            path: ['delegatedPermissions']
        });
    }
    async commandAction(logger, args) {
        try {
            const appObject = await this.getAppObject();
            const servicePrincipals = await odata.getAllItems(`${this.resource}/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);
            const appPermissions = [];
            if (args.options.delegatedPermissions) {
                const delegatedPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.delegatedPermissions, ScopeType.Scope, appPermissions, logger);
                this.addPermissionsToResourceArray(delegatedPermissions, appObject.requiredResourceAccess);
            }
            if (args.options.applicationPermissions) {
                const applicationPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.applicationPermissions, ScopeType.Role, appPermissions, logger);
                this.addPermissionsToResourceArray(applicationPermissions, appObject.requiredResourceAccess);
            }
            const addPermissionsRequestOptions = {
                url: `${this.resource}/v1.0/myorganization/applications/${appObject.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    requiredResourceAccess: appObject.requiredResourceAccess
                }
            };
            await request.patch(addPermissionsRequestOptions);
            if (args.options.grantAdminConsent) {
                let appServicePrincipal = servicePrincipals.find(sp => sp.appId === this.appId);
                if (!appServicePrincipal) {
                    if (this.verbose) {
                        await logger.logToStderr(`Creating service principal for app ${this.appId}...`);
                    }
                    appServicePrincipal = await entraServicePrincipal.createServicePrincipal(this.appId);
                }
                await this.grantAdminConsent(appServicePrincipal, appPermissions, logger);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppObject() {
        const apps = await odata.getAllItems(`${this.resource}/v1.0/myorganization/applications?$filter=appId eq '${formatting.encodeQueryParameter(this.appId)}'&$select=id,requiredResourceAccess`);
        if (apps.length === 0) {
            throw `App with id ${this.appId} not found in Microsoft Entra ID.`;
        }
        return apps[0];
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
    async grantOAuth2Permission(servicePricipalId, resourceId, scope) {
        const grantAdminConsentApplicationRequestOptions = {
            url: `${this.resource}/v1.0/myorganization/oauth2PermissionGrants`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                clientId: servicePricipalId,
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
            url: `${this.resource}/v1.0/myorganization/servicePrincipals/${servicePrincipalId}/appRoleAssignments`,
            headers: {
                'content-type': 'application/json'
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
}
export default new AppPermissionAddCommand();
//# sourceMappingURL=permission-add.js.map