var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppRoleAssignmentRemoveCommand_instances, _EntraAppRoleAssignmentRemoveCommand_initTelemetry, _EntraAppRoleAssignmentRemoveCommand_initOptions, _EntraAppRoleAssignmentRemoveCommand_initValidators, _EntraAppRoleAssignmentRemoveCommand_initOptionSets;
import os from 'os';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
class EntraAppRoleAssignmentRemoveCommand extends GraphCommand {
    get name() {
        return commands.APPROLEASSIGNMENT_REMOVE;
    }
    get description() {
        return 'Deletes an app role assignment for the specified Entra Application Registration';
    }
    constructor() {
        super();
        _EntraAppRoleAssignmentRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentRemoveCommand_instances, "m", _EntraAppRoleAssignmentRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentRemoveCommand_instances, "m", _EntraAppRoleAssignmentRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentRemoveCommand_instances, "m", _EntraAppRoleAssignmentRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAssignmentRemoveCommand_instances, "m", _EntraAppRoleAssignmentRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeAppRoleAssignment = async () => {
            let sp;
            // get the service principal associated with the appId
            let spMatchQuery;
            if (args.options.appId) {
                spMatchQuery = `appId eq '${formatting.encodeQueryParameter(args.options.appId)}'`;
            }
            else if (args.options.appObjectId) {
                spMatchQuery = `id eq '${formatting.encodeQueryParameter(args.options.appObjectId)}'`;
            }
            else {
                spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName)}'`;
            }
            try {
                let resp = await this.getServicePrincipalForApp(spMatchQuery);
                if (!resp.value.length) {
                    throw 'app registration not found';
                }
                sp = resp.value[0];
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
                resp = await request.get(requestOptions);
                if (!resp.value.length) {
                    throw 'Resource not found';
                }
                const appRolesToBeDeleted = [];
                const appRolesFound = resp.value[0].appRoles;
                if (!appRolesFound.length) {
                    throw `The resource '${args.options.resource}' does not have any application permissions available.`;
                }
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
                    appRolesToBeDeleted.push(existingRoles[0]);
                }
                const tasks = [];
                for (const appRole of appRolesToBeDeleted) {
                    const appRoleAssignment = sp.appRoleAssignments.filter((role) => role.appRoleId === appRole.id);
                    if (!appRoleAssignment.length) {
                        throw 'App role assignment not found';
                    }
                    tasks.push(this.removeAppRoleAssignmentForServicePrincipal(sp.id, appRoleAssignment[0].id));
                }
                await Promise.all(tasks);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeAppRoleAssignment();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the appRoleAssignment with scope(s) ${args.options.scopes} for resource ${args.options.resource}?` });
            if (result) {
                await removeAppRoleAssignment();
            }
        }
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
    async removeAppRoleAssignmentForServicePrincipal(spId, appRoleAssignmentId) {
        const spRequestOptions = {
            url: `${this.resource}/v1.0/servicePrincipals/${spId}/appRoleAssignments/${appRoleAssignmentId}`,
            headers: {
                'accept': 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.delete(spRequestOptions);
    }
}
_EntraAppRoleAssignmentRemoveCommand_instances = new WeakSet(), _EntraAppRoleAssignmentRemoveCommand_initTelemetry = function _EntraAppRoleAssignmentRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appDisplayName: typeof args.options.appDisplayName !== 'undefined',
            appObjectId: typeof args.options.appObjectId !== 'undefined',
            force: (!!args.options.force).toString()
        });
    });
}, _EntraAppRoleAssignmentRemoveCommand_initOptions = function _EntraAppRoleAssignmentRemoveCommand_initOptions() {
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
    }, {
        option: '-f, --force'
    });
}, _EntraAppRoleAssignmentRemoveCommand_initValidators = function _EntraAppRoleAssignmentRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
            return `${args.options.appObjectId} is not a valid GUID`;
        }
        return true;
    });
}, _EntraAppRoleAssignmentRemoveCommand_initOptionSets = function _EntraAppRoleAssignmentRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appObjectId', 'appDisplayName'] });
};
export default new EntraAppRoleAssignmentRemoveCommand();
//# sourceMappingURL=approleassignment-remove.js.map