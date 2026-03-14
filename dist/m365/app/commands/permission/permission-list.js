import request from '../../../../request.js';
import { entraApp } from '../../../../utils/entraApp.js';
import AppCommand, { appCommandOptions } from '../../../base/AppCommand.js';
import commands from '../../commands.js';
var GetServicePrincipal;
(function (GetServicePrincipal) {
    GetServicePrincipal[GetServicePrincipal["withPermissions"] = 0] = "withPermissions";
    GetServicePrincipal[GetServicePrincipal["withPermissionDefinitions"] = 1] = "withPermissionDefinitions";
})(GetServicePrincipal || (GetServicePrincipal = {}));
class AppPermissionListCommand extends AppCommand {
    get name() {
        return commands.PERMISSION_LIST;
    }
    get description() {
        return 'Lists API permissions for the current Microsoft Entra app';
    }
    get schema() {
        return appCommandOptions;
    }
    async commandAction(logger) {
        try {
            const servicePrincipal = await this.getServicePrincipal({ appId: this.appId }, logger, GetServicePrincipal.withPermissions);
            let permissions;
            if (servicePrincipal) {
                // service principal found, get permissions from the service principal
                permissions = await this.getServicePrincipalPermissions(servicePrincipal, logger);
            }
            else {
                // service principal not found, get permissions from app registration
                permissions = await this.getAppRegPermissions(this.appId, logger);
            }
            await logger.log(permissions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getServicePrincipal(servicePrincipalInfo, logger, mode) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving service principal ${servicePrincipalInfo.appId ?? servicePrincipalInfo.id}`);
        }
        const lookupUrl = servicePrincipalInfo.appId ? `?$filter=appId eq '${servicePrincipalInfo.appId}'&` : `/${servicePrincipalInfo.id}?`;
        const requestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals${lookupUrl}$select=appId,id,displayName`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if ((servicePrincipalInfo.id && !response) ||
            (servicePrincipalInfo.appId && response.value.length === 0)) {
            return undefined;
        }
        const servicePrincipal = servicePrincipalInfo.appId ?
            response.value[0] :
            response;
        if (this.verbose) {
            await logger.logToStderr(`Retrieving permissions for service principal ${servicePrincipal.id}...`);
        }
        const permissionsPromises = [];
        switch (mode) {
            case GetServicePrincipal.withPermissions: {
                const appRoleAssignmentsRequestOptions = {
                    url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/appRoleAssignments`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                const oauth2PermissionGrantsRequestOptions = {
                    url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/oauth2PermissionGrants`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                permissionsPromises.push(...[
                    request.get(appRoleAssignmentsRequestOptions),
                    request.get(oauth2PermissionGrantsRequestOptions)
                ]);
                break;
            }
            case GetServicePrincipal.withPermissionDefinitions: {
                const oauth2PermissionScopesRequestOptions = {
                    url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/oauth2PermissionScopes`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                const appRolesRequestOptions = {
                    url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/appRoles`,
                    headers: {
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                permissionsPromises.push(...[
                    request.get(oauth2PermissionScopesRequestOptions),
                    request.get(appRolesRequestOptions)
                ]);
                break;
            }
        }
        const permissions = await Promise.all(permissionsPromises);
        switch (mode) {
            case GetServicePrincipal.withPermissions:
                servicePrincipal.appRoleAssignments = permissions[0].value;
                servicePrincipal.oauth2PermissionGrants = permissions[1].value;
                break;
            case GetServicePrincipal.withPermissionDefinitions:
                servicePrincipal.oauth2PermissionScopes = permissions[0].value;
                servicePrincipal.appRoles = permissions[1].value;
                break;
        }
        return servicePrincipal;
    }
    async getServicePrincipalPermissions(servicePrincipal, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Resolving permissions for the service principal...`);
        }
        const apiPermissions = [];
        // hash table for resolving resource IDs to names
        const resourceLookup = {};
        // list of service principals for which to load permissions
        const servicePrincipalsToResolve = [];
        const appRoleAssignments = servicePrincipal.appRoleAssignments;
        apiPermissions.push(...appRoleAssignments.map(appRoleAssignment => {
            // store resource name for resolving OAuth2 grants
            resourceLookup[appRoleAssignment.resourceId] = appRoleAssignment.resourceDisplayName;
            // add to the list of service principals to load to get the app role
            // display name
            if (!servicePrincipalsToResolve.find(r => r.id === appRoleAssignment.resourceId)) {
                servicePrincipalsToResolve.push({ id: appRoleAssignment.resourceId });
            }
            return {
                resource: appRoleAssignment.resourceDisplayName,
                // we store the app role ID temporarily and will later resolve to display name
                permission: appRoleAssignment.appRoleId,
                type: 'Application'
            };
        }));
        const oauth2Grants = servicePrincipal.oauth2PermissionGrants;
        oauth2Grants.forEach(oauth2Grant => {
            // see if we can resolve the resource name from the resources
            // retrieved from app role assignments
            const resource = resourceLookup[oauth2Grant.resourceId] ?? oauth2Grant.resourceId;
            if (resource === oauth2Grant.resourceId &&
                !servicePrincipalsToResolve.find(r => r.id === oauth2Grant.resourceId)) {
                // resource name not found in the resources
                // add it to the list of resources to resolve
                servicePrincipalsToResolve.push({ id: oauth2Grant.resourceId });
            }
            const scopes = oauth2Grant.scope.split(' ');
            scopes.forEach(scope => {
                apiPermissions.push({
                    resource,
                    permission: scope,
                    type: 'Delegated'
                });
            });
        });
        if (servicePrincipalsToResolve.length > 0) {
            const servicePrincipals = await Promise
                .all(servicePrincipalsToResolve
                .map(servicePrincipalInfo => this.getServicePrincipal(servicePrincipalInfo, logger, GetServicePrincipal.withPermissionDefinitions)));
            servicePrincipals.forEach(servicePrincipal => {
                apiPermissions.forEach(apiPermission => {
                    if (apiPermission.resource === servicePrincipal.id) {
                        apiPermission.resource = servicePrincipal.displayName;
                    }
                    if (apiPermission.resource === servicePrincipal.displayName &&
                        apiPermission.type === 'Application') {
                        apiPermission.permission = servicePrincipal.appRoles
                            .find(appRole => appRole.id === apiPermission.permission)?.value ?? apiPermission.permission;
                    }
                });
            });
        }
        return apiPermissions;
    }
    async getAppRegistration(appId, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving the Entra application registration with appId '${appId}'`);
        }
        const app = await entraApp.getAppRegistrationByAppId(appId);
        return app;
    }
    async getAppRegPermissions(appId, logger) {
        const application = await this.getAppRegistration(appId, logger);
        if (application.requiredResourceAccess.length === 0) {
            return [];
        }
        const servicePrincipalsToResolve = application.requiredResourceAccess
            .map(resourceAccess => {
            return {
                appId: resourceAccess.resourceAppId
            };
        });
        const servicePrincipals = await Promise
            .all(servicePrincipalsToResolve.map(servicePrincipalInfo => this.getServicePrincipal(servicePrincipalInfo, logger, GetServicePrincipal.withPermissionDefinitions)));
        const apiPermissions = [];
        application.requiredResourceAccess.forEach(requiredResourceAccess => {
            const servicePrincipal = servicePrincipals
                .find(servicePrincipal => servicePrincipal?.appId === requiredResourceAccess.resourceAppId);
            const resourceName = servicePrincipal?.displayName ?? requiredResourceAccess.resourceAppId;
            requiredResourceAccess.resourceAccess.forEach(permission => {
                apiPermissions.push({
                    resource: resourceName,
                    permission: this.getPermissionName(permission.id, permission.type, servicePrincipal),
                    type: permission.type === 'Role' ? 'Application' : 'Delegated'
                });
            });
        });
        return apiPermissions;
    }
    getPermissionName(permissionId, permissionType, servicePrincipal) {
        if (!servicePrincipal) {
            return permissionId;
        }
        switch (permissionType) {
            case 'Role':
                return servicePrincipal.appRoles
                    .find(appRole => appRole.id === permissionId)?.value ?? permissionId;
            case 'Scope':
                return servicePrincipal.oauth2PermissionScopes
                    .find(permissionScope => permissionScope.id === permissionId)?.value ?? permissionId;
        }
        /* c8 ignore next 4 */
        // permissionType is either 'Scope' or 'Role' but we need a safe default
        // to avoid building errors. This code will never be reached.
        return permissionId;
    }
}
export default new AppPermissionListCommand();
//# sourceMappingURL=permission-list.js.map