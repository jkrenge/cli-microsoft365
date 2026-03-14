var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppRoleAssignmentAddCommand_instances, _EntraAppRoleAssignmentAddCommand_initTelemetry, _EntraAppRoleAssignmentAddCommand_initOptions, _EntraAppRoleAssignmentAddCommand_initValidators, _EntraAppRoleAssignmentAddCommand_initOptionSets;
import os from 'os';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
class EntraAppRoleAssignmentAddCommand extends GraphCommand {
    get name() {
        return commands.APPROLEASSIGNMENT_ADD;
    }
    get description() {
        return 'Adds service principal permissions also known as scopes and app role assignments for specified Microsoft Entra application registration';
    }
    constructor() {
        super();
        _EntraAppRoleAssignmentAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentAddCommand_instances, "m", _EntraAppRoleAssignmentAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentAddCommand_instances, "m", _EntraAppRoleAssignmentAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentAddCommand_instances, "m", _EntraAppRoleAssignmentAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentAddCommand_instances, "m", _EntraAppRoleAssignmentAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        let objectId;
        let queryFilter;
        if (args.options.appId) {
            queryFilter = `$filter=appId eq '${formatting.encodeQueryParameter(args.options.appId)}'`;
        }
        else if (args.options.appObjectId) {
            queryFilter = `$filter=id eq '${formatting.encodeQueryParameter(args.options.appObjectId)}'`;
        }
        else {
            queryFilter = `$filter=displayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName)}'`;
        }
        const getServicePrinciplesRequestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals?${queryFilter}`,
            headers: {
                accept: 'application/json'
            },
            responseType: 'json'
        };
        try {
            const servicePrincipalResult = await request.get(getServicePrinciplesRequestOptions);
            if (servicePrincipalResult.value.length === 0) {
                throw `The specified service principal doesn't exist`;
            }
            if (servicePrincipalResult.value.length > 1) {
                const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', servicePrincipalResult.value);
                const result = await cli.handleMultipleResultsFound(`Multiple service principal found.`, resultAsKeyValuePair);
                objectId = result.id;
            }
            else {
                objectId = servicePrincipalResult.value[0].id;
            }
            let resource = formatting.encodeQueryParameter(args.options.resource);
            // try resolve aliases that the user might enter since these are seen in the Azure portal
            switch (args.options.resource.toLocaleLowerCase()) {
                case 'sharepoint':
                    resource = 'Office 365 SharePoint Online';
                    break;
                case 'intune':
                    resource = 'Microsoft Intune API';
                    break;
                case 'exchange':
                    resource = 'Office 365 Exchange Online';
                    break;
            }
            // will perform resource name, appId or objectId search
            let filter = `$filter=(displayName eq '${resource}' or startswith(displayName,'${resource}'))`;
            if (validation.isValidGuid(resource)) {
                filter += ` or appId eq '${resource}' or id eq '${resource}'`;
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/servicePrincipals?${filter}`,
                headers: {
                    'accept': 'application/json'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            const appRoles = [];
            // flatten the app roles found
            const appRolesFound = [];
            for (const servicePrincipal of res.value) {
                for (const role of servicePrincipal.appRoles) {
                    appRolesFound.push({
                        resourceId: servicePrincipal.id,
                        objectId: role.id,
                        value: role.value
                    });
                }
            }
            if (!appRolesFound.length) {
                throw `The resource '${args.options.resource}' does not have any application permissions available.`;
            }
            // search for match between the found app roles and the specified scopes option value
            for (const scope of args.options.scopes.split(',')) {
                const existingRoles = appRolesFound.filter((role) => {
                    return role.value.toLocaleLowerCase() === scope.toLocaleLowerCase().trim();
                });
                if (!existingRoles.length) {
                    // the role specified in the scopes option does not belong to the found service principles
                    // throw an error and show list with available roles (scopes)
                    let availableRoles = '';
                    appRolesFound.map((r) => availableRoles += `${os.EOL}${r.value}`);
                    throw `The scope value '${scope}' you have specified does not exist for ${args.options.resource}. ${os.EOL}Available scopes (application permissions) are: ${availableRoles}`;
                }
                appRoles.push(existingRoles[0]);
            }
            const tasks = [];
            for (const appRole of appRoles) {
                tasks.push(this.addRoleToServicePrincipal(objectId, appRole));
            }
            const rolesAddedResponse = await Promise.all(tasks);
            if (args.options.output && args.options.output.toLowerCase() === 'json') {
                await logger.log(rolesAddedResponse);
            }
            else {
                await logger.log(rolesAddedResponse.map((result) => ({
                    objectId: result.id,
                    principalDisplayName: result.principalDisplayName,
                    resourceDisplayName: result.resourceDisplayName
                })));
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async addRoleToServicePrincipal(objectId, appRole) {
        const requestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals/${objectId}/appRoleAssignments`,
            headers: {
                'Content-Type': 'application/json'
            },
            responseType: 'json',
            data: {
                appRoleId: appRole.objectId,
                principalId: objectId,
                resourceId: appRole.resourceId
            }
        };
        return request.post(requestOptions);
    }
}
_EntraAppRoleAssignmentAddCommand_instances = new WeakSet(), _EntraAppRoleAssignmentAddCommand_initTelemetry = function _EntraAppRoleAssignmentAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appObjectId: typeof args.options.appObjectId !== 'undefined',
            appDisplayName: typeof args.options.appDisplayName !== 'undefined'
        });
    });
}, _EntraAppRoleAssignmentAddCommand_initOptions = function _EntraAppRoleAssignmentAddCommand_initOptions() {
    this.options.unshift({
        option: '--appId [appId]'
    }, {
        option: '--appObjectId [appObjectId]'
    }, {
        option: '--appDisplayName [appDisplayName]'
    }, {
        option: '-r, --resource <resource>',
        autocomplete: ['Microsoft Graph', 'SharePoint', 'OneNote', 'Exchange', 'Microsoft Forms', 'Azure Active Directory Graph', 'Skype for Business']
    }, {
        option: '-s, --scopes <scopes>'
    });
}, _EntraAppRoleAssignmentAddCommand_initValidators = function _EntraAppRoleAssignmentAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
            return `${args.options.appObjectId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraAppRoleAssignmentAddCommand_initOptionSets = function _EntraAppRoleAssignmentAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appObjectId', 'appDisplayName'] });
};
export default new EntraAppRoleAssignmentAddCommand();
//# sourceMappingURL=approleassignment-add.js.map