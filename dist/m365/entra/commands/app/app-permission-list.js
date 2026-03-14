var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppPermissionListCommand_instances, _EntraAppPermissionListCommand_initTelemetry, _EntraAppPermissionListCommand_initOptions, _EntraAppPermissionListCommand_initValidators, _EntraAppPermissionListCommand_initOptionSets;
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import request from "../../../../request.js";
import { validation } from "../../../../utils/validation.js";
import { entraApp } from "../../../../utils/entraApp.js";
class EntraAppPermissionListCommand extends GraphCommand {
    get name() {
        return commands.APP_PERMISSION_LIST;
    }
    get description() {
        return 'Lists the application and delegated permissions for a specified Entra Application Registration';
    }
    constructor() {
        super();
        _EntraAppPermissionListCommand_instances.add(this);
        this.allowedTypes = ['delegated', 'application', 'all'];
        __classPrivateFieldGet(this, _EntraAppPermissionListCommand_instances, "m", _EntraAppPermissionListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionListCommand_instances, "m", _EntraAppPermissionListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionListCommand_instances, "m", _EntraAppPermissionListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppPermissionListCommand_instances, "m", _EntraAppPermissionListCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appObjectId = await this.getAppObjectId(args.options, logger);
            const type = args.options.type ?? 'all';
            const permissions = await this.getAppRegPermissions(appObjectId, type, logger);
            await logger.log(permissions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppObjectId(options, logger) {
        if (options.appObjectId) {
            return options.appObjectId;
        }
        const { appId, appName } = options;
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : appName}...`);
        }
        if (appId) {
            const app = await entraApp.getAppRegistrationByAppId(appId, ["id"]);
            return app.id;
        }
        else {
            const app = await entraApp.getAppRegistrationByAppName(appName, ["id"]);
            return app.id;
        }
    }
    async getAppRegPermissions(appObjectId, permissionType, logger) {
        const requestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${appObjectId}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const application = await request.get(requestOptions);
        const requiredResourceAccess = application.requiredResourceAccess;
        if (requiredResourceAccess.length === 0) {
            return [];
        }
        const servicePrincipalsToResolve = requiredResourceAccess.map(resourceAccess => {
            return {
                appId: resourceAccess.resourceAppId
            };
        });
        const servicePrincipals = await Promise
            .all(servicePrincipalsToResolve.map(servicePrincipalInfo => this.getServicePrincipal(servicePrincipalInfo, permissionType, logger)));
        const apiPermissions = [];
        requiredResourceAccess.forEach(requiredResourceAccess => {
            const servicePrincipal = servicePrincipals
                .find(servicePrincipal => servicePrincipal?.appId === requiredResourceAccess.resourceAppId);
            const resourceName = servicePrincipal?.displayName ?? requiredResourceAccess.resourceAppId;
            requiredResourceAccess.resourceAccess.forEach(permission => {
                if (permissionType === 'application' && permission.type === 'Scope') {
                    return;
                }
                if (permissionType === 'delegated' && permission.type === 'Role') {
                    return;
                }
                apiPermissions.push({
                    resource: resourceName,
                    resourceId: requiredResourceAccess.resourceAppId,
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
        if (permissionType === 'Role') {
            return servicePrincipal.appRoles
                .find(appRole => appRole.id === permissionId)?.value ?? permissionId;
        }
        // permissionType === 'Scope'
        return servicePrincipal.oauth2PermissionScopes
            .find(permissionScope => permissionScope.id === permissionId)?.value ?? permissionId;
    }
    async getServicePrincipal(servicePrincipalInfo, permissionType, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving service principal ${servicePrincipalInfo.appId}`);
        }
        const requestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals?$filter=appId eq '${servicePrincipalInfo.appId}'&$select=appId,id,displayName`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        if (servicePrincipalInfo.appId && response.value.length === 0) {
            return null;
        }
        const servicePrincipal = response.value[0];
        if (this.verbose) {
            await logger.logToStderr(`Retrieving permissions for service principal ${servicePrincipal.id}...`);
        }
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
        let permissions;
        if (permissionType === 'all' || permissionType === 'delegated') {
            permissions = await request.get(oauth2PermissionScopesRequestOptions);
            servicePrincipal.oauth2PermissionScopes = permissions.value;
        }
        if (permissionType === 'all' || permissionType === 'application') {
            permissions = await request.get(appRolesRequestOptions);
            servicePrincipal.appRoles = permissions.value;
        }
        return servicePrincipal;
    }
}
_EntraAppPermissionListCommand_instances = new WeakSet(), _EntraAppPermissionListCommand_initTelemetry = function _EntraAppPermissionListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appName: typeof args.options.appName !== 'undefined',
            appObjectId: typeof args.options.appObjectId !== 'undefined',
            type: typeof args.options.type !== 'undefined'
        });
    });
}, _EntraAppPermissionListCommand_initOptions = function _EntraAppPermissionListCommand_initOptions() {
    this.options.unshift({ option: '-i, --appId [appId]' }, { option: '-n, --appName [appName]' }, { option: '--appObjectId [appObjectId]' }, { option: '--type [type]', autocomplete: this.allowedTypes });
}, _EntraAppPermissionListCommand_initValidators = function _EntraAppPermissionListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
            return `${args.options.appObjectId} is not a valid GUID`;
        }
        if (args.options.type && this.allowedTypes.indexOf(args.options.type.toLowerCase()) === -1) {
            return `${args.options.type} is not a valid type. Allowed types are ${this.allowedTypes.join(', ')}`;
        }
        return true;
    });
}, _EntraAppPermissionListCommand_initOptionSets = function _EntraAppPermissionListCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appName', 'appObjectId'] });
};
export default new EntraAppPermissionListCommand();
//# sourceMappingURL=app-permission-list.js.map