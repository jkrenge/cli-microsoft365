var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFeatureListCommand_instances, _SpoFeatureListCommand_initTelemetry, _SpoFeatureListCommand_initOptions, _SpoFeatureListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFeatureListCommand extends SpoCommand {
    get name() {
        return commands.FEATURE_LIST;
    }
    get description() {
        return 'Lists Features activated in the specified site or site collection';
    }
    constructor() {
        super();
        _SpoFeatureListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFeatureListCommand_instances, "m", _SpoFeatureListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFeatureListCommand_instances, "m", _SpoFeatureListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFeatureListCommand_instances, "m", _SpoFeatureListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const scope = (args.options.scope) ? args.options.scope : 'Web';
        try {
            const features = await odata.getAllItems(`${args.options.webUrl}/_api/${scope}/Features?$select=DisplayName,DefinitionId`);
            await logger.log(features);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFeatureListCommand_instances = new WeakSet(), _SpoFeatureListCommand_initTelemetry = function _SpoFeatureListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            scope: args.options.scope || 'Web'
        });
    });
}, _SpoFeatureListCommand_initOptions = function _SpoFeatureListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web']
    });
}, _SpoFeatureListCommand_initValidators = function _SpoFeatureListCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.scope) {
            if (args.options.scope !== 'Site' &&
                args.options.scope !== 'Web') {
                return `${args.options.scope} is not a valid Feature scope. Allowed values are Site|Web`;
            }
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
};
export default new SpoFeatureListCommand();
//# sourceMappingURL=feature-list.js.map