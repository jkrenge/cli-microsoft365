var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFeatureEnableCommand_instances, _SpoFeatureEnableCommand_initTelemetry, _SpoFeatureEnableCommand_initOptions, _SpoFeatureEnableCommand_initValidators, _SpoFeatureEnableCommand_initTypes;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFeatureEnableCommand extends SpoCommand {
    get name() {
        return commands.FEATURE_ENABLE;
    }
    get description() {
        return 'Enables feature for the specified site or web';
    }
    constructor() {
        super();
        _SpoFeatureEnableCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFeatureEnableCommand_instances, "m", _SpoFeatureEnableCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFeatureEnableCommand_instances, "m", _SpoFeatureEnableCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFeatureEnableCommand_instances, "m", _SpoFeatureEnableCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFeatureEnableCommand_instances, "m", _SpoFeatureEnableCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        let scope = args.options.scope;
        const force = !!args.options.force;
        if (!scope) {
            scope = "web";
        }
        if (this.verbose) {
            await logger.logToStderr(`Enabling feature '${args.options.id}' on scope '${scope}' for url '${args.options.webUrl}' (force='${force}')...`);
        }
        const url = `${args.options.webUrl}/_api/${scope}/features/add(featureId=guid'${args.options.id}',force=${force})`;
        const requestOptions = {
            url: url,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFeatureEnableCommand_instances = new WeakSet(), _SpoFeatureEnableCommand_initTelemetry = function _SpoFeatureEnableCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            scope: args.options.scope || 'web',
            force: args.options.force || false
        });
    });
}, _SpoFeatureEnableCommand_initOptions = function _SpoFeatureEnableCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --id <id>'
    }, {
        option: '-s, --scope [scope]',
        autocomplete: ['Site', 'Web']
    }, {
        option: '--force'
    });
}, _SpoFeatureEnableCommand_initValidators = function _SpoFeatureEnableCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID for option id.`;
        }
        if (args.options.scope) {
            if (['site', 'web'].indexOf(args.options.scope.toLowerCase()) < 0) {
                return `${args.options.scope} is not a valid Feature scope. Allowed values are Site|Web`;
            }
        }
        return validation.isValidSharePointUrl(args.options.webUrl);
    });
}, _SpoFeatureEnableCommand_initTypes = function _SpoFeatureEnableCommand_initTypes() {
    this.types.string.push('scope', 's');
};
export default new SpoFeatureEnableCommand();
//# sourceMappingURL=feature-enable.js.map