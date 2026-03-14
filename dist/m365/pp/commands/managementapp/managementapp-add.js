var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpManagementAppAddCommand_instances, _PpManagementAppAddCommand_initTelemetry, _PpManagementAppAddCommand_initOptions, _PpManagementAppAddCommand_initValidators, _PpManagementAppAddCommand_initOptionSets;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { entraApp } from '../../../../utils/entraApp.js';
class PpManagementAppAddCommand extends PowerPlatformCommand {
    get name() {
        return commands.MANAGEMENTAPP_ADD;
    }
    get description() {
        return 'Register management application for Power Platform';
    }
    constructor() {
        super();
        _PpManagementAppAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpManagementAppAddCommand_instances, "m", _PpManagementAppAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpManagementAppAddCommand_instances, "m", _PpManagementAppAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PpManagementAppAddCommand_instances, "m", _PpManagementAppAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _PpManagementAppAddCommand_instances, "m", _PpManagementAppAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appId = await this.getAppId(args);
            const requestOptions = {
                // This should be refactored once we implement a PowerPlatform base class as api.bap will differ between envs.
                url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/adminApplications/${appId}?api-version=2020-06-01`,
                headers: {
                    accept: 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const res = await request.put(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppId(args) {
        if (args.options.appId) {
            return args.options.appId;
        }
        const { objectId, name } = args.options;
        if (objectId) {
            const app = await entraApp.getAppRegistrationByObjectId(objectId, ['appId']);
            return app.appId;
        }
        else {
            const app = await entraApp.getAppRegistrationByAppName(name, ['appId']);
            return app.appId;
        }
    }
}
_PpManagementAppAddCommand_instances = new WeakSet(), _PpManagementAppAddCommand_initTelemetry = function _PpManagementAppAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appId: typeof args.options.appId !== 'undefined',
            objectId: typeof args.options.objectId !== 'undefined',
            name: typeof args.options.name !== 'undefined'
        });
    });
}, _PpManagementAppAddCommand_initOptions = function _PpManagementAppAddCommand_initOptions() {
    this.options.unshift({ option: '--appId [appId]' }, { option: '--objectId [objectId]' }, { option: '--name [name]' });
}, _PpManagementAppAddCommand_initValidators = function _PpManagementAppAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
            return `${args.options.appId} is not a valid GUID`;
        }
        if (args.options.objectId && !validation.isValidGuid(args.options.objectId)) {
            return `${args.options.objectId} is not a valid GUID`;
        }
        return true;
    });
}, _PpManagementAppAddCommand_initOptionSets = function _PpManagementAppAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['appId', 'objectId', 'name'] });
};
export default new PpManagementAppAddCommand();
//# sourceMappingURL=managementapp-add.js.map