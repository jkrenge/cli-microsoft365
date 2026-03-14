var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppRoleAddCommand_instances, _a, _EntraAppRoleAddCommand_initTelemetry, _EntraAppRoleAddCommand_initOptions, _EntraAppRoleAddCommand_initValidators, _EntraAppRoleAddCommand_initOptionSets;
import { v4 } from 'uuid';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraApp } from '../../../../utils/entraApp.js';
class EntraAppRoleAddCommand extends GraphCommand {
    get name() {
        return commands.APP_ROLE_ADD;
    }
    get description() {
        return 'Adds role to the specified Entra app registration';
    }
    constructor() {
        super();
        _EntraAppRoleAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppRoleAddCommand_instances, "m", _EntraAppRoleAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAddCommand_instances, "m", _EntraAppRoleAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAddCommand_instances, "m", _EntraAppRoleAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleAddCommand_instances, "m", _EntraAppRoleAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appInfo = await this.getAppInfo(args, logger);
            if (this.verbose) {
                await logger.logToStderr(`Adding role ${args.options.name} to Microsoft Entra app ${appInfo.id}...`);
            }
            const requestOptions = {
                url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json',
                data: {
                    appRoles: appInfo.appRoles.concat({
                        displayName: args.options.name,
                        description: args.options.description,
                        id: v4(),
                        value: args.options.claim,
                        allowedMemberTypes: this.getAllowedMemberTypes(args)
                    })
                }
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getAllowedMemberTypes(args) {
        switch (args.options.allowedMembers) {
            case 'usersGroups':
                return ['User'];
            case 'applications':
                return ['Application'];
            case 'both':
                return ['User', 'Application'];
            default:
                return [];
        }
    }
    async getAppInfo(args, logger) {
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
_a = EntraAppRoleAddCommand, _EntraAppRoleAddCommand_instances = new WeakSet(), _EntraAppRoleAddCommand_initTelemetry = function _EntraAppRoleAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appObjectId: typeof args.options.appObjectId !== 'undefined',
            appName: typeof args.options.appName !== 'undefined'
        });
    });
}, _EntraAppRoleAddCommand_initOptions = function _EntraAppRoleAddCommand_initOptions() {
    this.options.unshift({ option: '--appId [appId]' }, { option: '--appObjectId [appObjectId]' }, { option: '--appName [appName]' }, { option: '-n, --name <name>' }, { option: '-d, --description <description>' }, {
        option: '-m, --allowedMembers <allowedMembers>', autocomplete: _a.allowedMembers
    }, { option: '-c, --claim <claim>' });
}, _EntraAppRoleAddCommand_initValidators = function _EntraAppRoleAddCommand_initValidators() {
    this.validators.push(async (args) => {
        const { allowedMembers, claim } = args.options;
        if (_a.allowedMembers.indexOf(allowedMembers) < 0) {
            return `${allowedMembers} is not a valid value for allowedMembers. Valid values are ${_a.allowedMembers.join(', ')}`;
        }
        if (claim.length > 120) {
            return `Claim must not be longer than 120 characters`;
        }
        if (claim.startsWith('.')) {
            return 'Claim must not begin with .';
        }
        if (!/^[\w:!#$%&'()*+,-./:;<=>?@[\]^+_`{|}~]+$/.test(claim)) {
            return `Claim can contain only the following characters a-z, A-Z, 0-9, :!#$%&'()*+,-./:;<=>?@[]^+_\`{|}~]+`;
        }
        return true;
    });
}, _EntraAppRoleAddCommand_initOptionSets = function _EntraAppRoleAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appObjectId', 'appName'] });
};
EntraAppRoleAddCommand.allowedMembers = ['usersGroups', 'applications', 'both'];
export default new EntraAppRoleAddCommand();
//# sourceMappingURL=app-role-add.js.map