var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _PpAiBuilderModelListCommand_instances, _PpAiBuilderModelListCommand_initTelemetry, _PpAiBuilderModelListCommand_initOptions;
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
class PpAiBuilderModelListCommand extends PowerPlatformCommand {
    get name() {
        return commands.AIBUILDERMODEL_LIST;
    }
    get description() {
        return 'List available AI builder models in the specified Power Platform environment.';
    }
    defaultProperties() {
        return ['msdyn_name', 'msdyn_aimodelid', 'createdon', 'modifiedon'];
    }
    constructor() {
        super();
        _PpAiBuilderModelListCommand_instances.add(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelListCommand_instances, "m", _PpAiBuilderModelListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _PpAiBuilderModelListCommand_instances, "m", _PpAiBuilderModelListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving available AI Builder models`);
        }
        try {
            const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
            const aimodels = await odata.getAllItems(`${dynamicsApiUrl}/api/data/v9.0/msdyn_aimodels?$filter=iscustomizable/Value eq true`);
            await logger.log(aimodels);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_PpAiBuilderModelListCommand_instances = new WeakSet(), _PpAiBuilderModelListCommand_initTelemetry = function _PpAiBuilderModelListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            asAdmin: !!args.options.asAdmin
        });
    });
}, _PpAiBuilderModelListCommand_initOptions = function _PpAiBuilderModelListCommand_initOptions() {
    this.options.unshift({
        option: '-e, --environmentName <environmentName>'
    }, {
        option: '--asAdmin'
    });
};
export default new PpAiBuilderModelListCommand();
//# sourceMappingURL=aibuildermodel-list.js.map