var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppRoleRemoveCommand_instances, _EntraAppRoleRemoveCommand_initTelemetry, _EntraAppRoleRemoveCommand_initOptions, _EntraAppRoleRemoveCommand_initValidators, _EntraAppRoleRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from "../../../../utils/formatting.js";
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraApp } from "../../../../utils/entraApp.js";
class EntraAppRoleRemoveCommand extends GraphCommand {
    get name() {
        return commands.APP_ROLE_REMOVE;
    }
    get description() {
        return 'Removes role from the specified Entra app registration';
    }
    constructor() {
        super();
        _EntraAppRoleRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppRoleRemoveCommand_instances, "m", _EntraAppRoleRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleRemoveCommand_instances, "m", _EntraAppRoleRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleRemoveCommand_instances, "m", _EntraAppRoleRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleRemoveCommand_instances, "m", _EntraAppRoleRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const deleteAppRole = async () => {
            try {
                await this.processAppRoleDelete(logger, args);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await deleteAppRole();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app role?` });
            if (result) {
                await deleteAppRole();
            }
        }
    }
    async processAppRoleDelete(logger, args) {
        const app = await this.getEntraApp(args, logger);
        const appRoleDeleteIdentifierNameValue = args.options.name ? `name '${args.options.name}'` : (args.options.claim ? `claim '${args.options.claim}'` : `id '${args.options.id}'`);
        if (this.verbose) {
            await logger.logToStderr(`Deleting role with ${appRoleDeleteIdentifierNameValue} from Microsoft Entra app ${app.id}...`);
        }
        // Find the role search criteria provided by the user.
        const appRoleDeleteIdentifierProperty = args.options.name ? `displayName` : (args.options.claim ? `value` : `id`);
        const appRoleDeleteIdentifierValue = args.options.name ? args.options.name : (args.options.claim ? args.options.claim : args.options.id);
        const appRoleToDelete = app.appRoles.filter((role) => role[appRoleDeleteIdentifierProperty] === appRoleDeleteIdentifierValue);
        if (args.options.name &&
            appRoleToDelete !== undefined &&
            appRoleToDelete.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', appRoleToDelete);
            appRoleToDelete[0] = await cli.handleMultipleResultsFound(`Multiple roles with name '${args.options.name}' found.`, resultAsKeyValuePair);
        }
        if (appRoleToDelete.length === 0) {
            throw `No app role with ${appRoleDeleteIdentifierNameValue} found.`;
        }
        const roleToDelete = appRoleToDelete[0];
        if (roleToDelete.isEnabled) {
            await this.disableAppRole(logger, app, roleToDelete.id);
            await this.deleteAppRole(logger, app, roleToDelete.id);
        }
        else {
            await this.deleteAppRole(logger, app, roleToDelete.id);
        }
    }
    async disableAppRole(logger, app, roleId) {
        const roleIndex = app.appRoles.findIndex((role) => role.id === roleId);
        if (this.verbose) {
            await logger.logToStderr(`Disabling the app role`);
        }
        app.appRoles[roleIndex].isEnabled = false;
        const requestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${app.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                appRoles: app.appRoles
            }
        };
        return request.patch(requestOptions);
    }
    async deleteAppRole(logger, app, roleId) {
        if (this.verbose) {
            await logger.logToStderr(`Deleting the app role.`);
        }
        const updatedAppRoles = app.appRoles.filter((role) => role.id !== roleId);
        const requestOptions = {
            url: `${this.resource}/v1.0/myorganization/applications/${app.id}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
                appRoles: updatedAppRoles
            }
        };
        return request.patch(requestOptions);
    }
    async getEntraApp(args, logger) {
        const { appObjectId, appId, appName } = args.options;
        if (this.verbose) {
            await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appObjectId ? appObjectId : (appId ? appId : appName)}...`);
        }
        if (appObjectId) {
            return await entraApp.getAppRegistrationByObjectId(appObjectId, ['id', 'appRoles']);
        }
        else if (appId) {
            return await entraApp.getAppRegistrationByAppId(appId, ['id', 'appRoles']);
        }
        else {
            return await entraApp.getAppRegistrationByAppName(appName, ['id', 'appRoles']);
        }
    }
}
_EntraAppRoleRemoveCommand_instances = new WeakSet(), _EntraAppRoleRemoveCommand_initTelemetry = function _EntraAppRoleRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appObjectId: typeof args.options.appObjectId !== 'undefined',
            appName: typeof args.options.appName !== 'undefined',
            claim: typeof args.options.claim !== 'undefined',
            name: typeof args.options.name !== 'undefined',
            id: typeof args.options.id !== 'undefined'
        });
    });
}, _EntraAppRoleRemoveCommand_initOptions = function _EntraAppRoleRemoveCommand_initOptions() {
    this.options.unshift({ option: '--appId [appId]' }, { option: '--appObjectId [appObjectId]' }, { option: '--appName [appName]' }, { option: '-n, --name [name]' }, { option: '-i, --id [id]' }, { option: '-c, --claim [claim]' }, { option: '-f, --force' });
}, _EntraAppRoleRemoveCommand_initValidators = function _EntraAppRoleRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id) {
            if (!validation.isValidGuid(args.options.id)) {
                return `${args.options.id} is not a valid GUID`;
            }
        }
        return true;
    });
}, _EntraAppRoleRemoveCommand_initOptionSets = function _EntraAppRoleRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appObjectId', 'appName'] }, { options: ['name', 'claim', 'id'] });
};
export default new EntraAppRoleRemoveCommand();
//# sourceMappingURL=app-role-remove.js.map