var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PaAppListCommand_instances, _PaAppListCommand_initTelemetry, _PaAppListCommand_initOptions, _PaAppListCommand_initValidators;
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
class PaAppListCommand extends PowerAppsCommand {
    get name() {
        return commands.APP_LIST;
    }
    get description() {
        return 'Lists all Power Apps apps';
    }
    defaultProperties() {
        return ['name', 'displayName'];
    }
    constructor() {
        super();
        _PaAppListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PaAppListCommand_instances, "m", _PaAppListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PaAppListCommand_instances, "m", _PaAppListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _PaAppListCommand_instances, "m", _PaAppListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const url = `${this.resource}/providers/Microsoft.PowerApps${args.options.asAdmin ? '/scopes/admin' : ''}${args.options.environmentName ? '/environments/' + formatting.encodeQueryParameter(args.options.environmentName) : ''}/apps?api-version=2017-08-01`;
        try {
            const apps = await odata.getAllItems(url);
            if (apps.length > 0) {
                apps.forEach(a => {
                    a.displayName = a.properties.displayName;
                });
            }
            await logger.log(apps);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PaAppListCommand_instances = new WeakSet(), _PaAppListCommand_initTelemetry = function _PaAppListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: args.options.asAdmin === true,
            environmentName: typeof args.options.environmentName !== 'undefined'
        });
    });
}, _PaAppListCommand_initOptions = function _PaAppListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName [environmentName]'
    }, {
        option: '--asAdmin'
    });
}, _PaAppListCommand_initValidators = function _PaAppListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.asAdmin && !args.options.environmentName) {
            return 'When specifying the asAdmin option the environment option is required as well';
        }
        if (args.options.environmentName && !args.options.asAdmin) {
            return 'When specifying the environment option the asAdmin option is required as well';
        }
        return true;
    });
};
export default new PaAppListCommand();
//# sourceMappingURL=app-list.js.map