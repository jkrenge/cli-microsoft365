var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppRoleAssignmentListCommand_instances, _EntraAppRoleAssignmentListCommand_initTelemetry, _EntraAppRoleAssignmentListCommand_initOptions, _EntraAppRoleAssignmentListCommand_initValidators, _EntraAppRoleAssignmentListCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraAppRoleAssignmentListCommand extends GraphCommand {
    get name() {
        return commands.APPROLEASSIGNMENT_LIST;
    }
    get description() {
        return 'Lists app role assignments for the specified application registration';
    }
    constructor() {
        super();
        _EntraAppRoleAssignmentListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentListCommand_instances, "m", _EntraAppRoleAssignmentListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentListCommand_instances, "m", _EntraAppRoleAssignmentListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentListCommand_instances, "m", _EntraAppRoleAssignmentListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentListCommand_instances, "m", _EntraAppRoleAssignmentListCommand_initOptionSets).call(this);
    }
    defaultProperties() {
        return ['resourceDisplayName', 'roleName'];
    }
    async commandAction(logger, args) {
        try {
            const spAppRoleAssignments = await this.getAppRoleAssignments(args.options);
            // the role assignment has an appRoleId but no name. To get the name,
            // we need to get all the roles from the resource. the resource is
            // a service principal. Multiple roles may have same resource id.
            const resourceIds = spAppRoleAssignments.map((item) => item.resourceId);
            const tasks = [];
            for (let i = 0; i < resourceIds.length; i++) {
                tasks.push(this.getServicePrincipal(resourceIds[i]));
            }
            const resources = await Promise.all(tasks);
            // loop through all appRoleAssignments for the servicePrincipal
            // and lookup the appRole.Id in the resources[resourceId].appRoles array...
            const results = [];
            spAppRoleAssignments.map((appRoleAssignment) => {
                const resource = resources.find((r) => r.id === appRoleAssignment.resourceId);
                if (resource) {
                    const appRole = resource.appRoles.find((r) => r.id === appRoleAssignment.appRoleId);
                    if (appRole) {
                        results.push({
                            appRoleId: appRoleAssignment.appRoleId,
                            resourceDisplayName: appRoleAssignment.resourceDisplayName,
                            resourceId: appRoleAssignment.resourceId,
                            roleId: appRole.id,
                            roleName: appRole.value,
                            created: appRoleAssignment.createdDateTime,
                            deleted: appRoleAssignment.deletedDateTime
                        });
                    }
                }
            });
            await logger.log(results);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppRoleAssignments(argOptions) {
        if (argOptions.appObjectId) {
            const spAppRoleAssignments = await this.getSPAppRoleAssignments(argOptions.appObjectId);
            if (!spAppRoleAssignments.value.length) {
                throw 'no app role assignments found';
            }
            return spAppRoleAssignments.value;
        }
        else {
            const spMatchQuery = argOptions.appId
                ? `appId eq '${formatting.encodeQueryParameter(argOptions.appId)}'`
                : `displayName eq '${formatting.encodeQueryParameter(argOptions.appDisplayName)}'`;
            const resp = await this.getServicePrincipalForApp(spMatchQuery);
            if (!resp.value.length) {
                throw 'app registration not found';
            }
            return resp.value[0].appRoleAssignments;
        }
    }
    async getSPAppRoleAssignments(spId) {
        const spRequestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals/${spId}/appRoleAssignments`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        return request.get(spRequestOptions);
    }
    async getServicePrincipalForApp(filterParam) {
        const spRequestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals?$expand=appRoleAssignments&$filter=${filterParam}`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        return request.get(spRequestOptions);
    }
    async getServicePrincipal(spId) {
        const spRequestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals/${spId}`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        return request.get(spRequestOptions);
    }
}
_EntraAppRoleAssignmentListCommand_instances = new WeakSet(), _EntraAppRoleAssignmentListCommand_initTelemetry = function _EntraAppRoleAssignmentListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appDisplayName: typeof args.options.appDisplayName !== 'undefined',
            appObjectId: typeof args.options.appObjectId !== 'undefined'
        });
    });
}, _EntraAppRoleAssignmentListCommand_initOptions = function _EntraAppRoleAssignmentListCommand_initOptions() {
    this.options.unshift({
        option: '-i, --appId [appId]'
    }, {
        option: '-n, --appDisplayName [appDisplayName]'
    }, {
        option: '--appObjectId [appObjectId]'
    });
}, _EntraAppRoleAssignmentListCommand_initValidators = function _EntraAppRoleAssignmentListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
            return `${args.options.appObjectId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraAppRoleAssignmentListCommand_initOptionSets = function _EntraAppRoleAssignmentListCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appObjectId', 'appDisplayName'] });
};
export default new EntraAppRoleAssignmentListCommand();
//# sourceMappingURL=approleassignment-list.js.map