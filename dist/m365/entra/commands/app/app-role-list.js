var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _EntraAppRoleListCommand_instances, _EntraAppRoleListCommand_initTelemetry, _EntraAppRoleListCommand_initOptions, _EntraAppRoleListCommand_initOptionSets;
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraApp } from '../../../../utils/entraApp.js';
class EntraAppRoleListCommand extends GraphCommand {
    get name() {
        return commands.APP_ROLE_LIST;
    }
    get description() {
        return 'Gets Entra app registration roles';
    }
    constructor() {
        super();
        _EntraAppRoleListCommand_instances.add(this);
        __classPrivateFieldGet(this, _EntraAppRoleListCommand_instances, "m", _EntraAppRoleListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleListCommand_instances, "m", _EntraAppRoleListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _EntraAppRoleListCommand_instances, "m", _EntraAppRoleListCommand_initOptionSets).call(this);
    }
    defaultProperties() {
        return ['displayName', 'description', 'id'];
    }
    async commandAction(logger, args) {
        try {
            const objectId = await this.getAppObjectId(args, logger);
            const appRoles = await odata.getAllItems(`${this.resource}/v1.0/myorganization/applications/${objectId}/appRoles`);
            await logger.log(appRoles);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppObjectId(args, logger) {
        if (args.options.appObjectId) {
            return args.options.appObjectId;
        }
        const { appId, appName } = args.options;
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
}
_EntraAppRoleListCommand_instances = new WeakSet(), _EntraAppRoleListCommand_initTelemetry = function _EntraAppRoleListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            appObjectId: typeof args.options.appObjectId !== 'undefined',
            appName: typeof args.options.appName !== 'undefined'
        });
    });
}, _EntraAppRoleListCommand_initOptions = function _EntraAppRoleListCommand_initOptions() {
    this.options.unshift({ option: '--appId [appId]' }, { option: '--appObjectId [appObjectId]' }, { option: '--appName [appName]' });
}, _EntraAppRoleListCommand_initOptionSets = function _EntraAppRoleListCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'appObjectId', 'appName'] });
};
export default new EntraAppRoleListCommand();
//# sourceMappingURL=app-role-list.js.map