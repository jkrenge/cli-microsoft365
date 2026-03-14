var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebInstalledLanguageListCommand_instances, _SpoWebInstalledLanguageListCommand_initOptions, _SpoWebInstalledLanguageListCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoWebInstalledLanguageListCommand extends SpoCommand {
    get name() {
        return commands.WEB_INSTALLEDLANGUAGE_LIST;
    }
    get description() {
        return 'Lists all installed languages on site';
    }
    defaultProperties() {
        return ['DisplayName', 'LanguageTag', 'Lcid'];
    }
    constructor() {
        super();
        _SpoWebInstalledLanguageListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebInstalledLanguageListCommand_instances, "m", _SpoWebInstalledLanguageListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebInstalledLanguageListCommand_instances, "m", _SpoWebInstalledLanguageListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving all installed languages on site ${args.options.webUrl}...`);
        }
        const requestOptions = {
            url: `${args.options.webUrl}/_api/web/RegionalSettings/InstalledLanguages`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const webInstalledLanguageProperties = await request.get(requestOptions);
            await logger.log(webInstalledLanguageProperties.Items);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoWebInstalledLanguageListCommand_instances = new WeakSet(), _SpoWebInstalledLanguageListCommand_initOptions = function _SpoWebInstalledLanguageListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoWebInstalledLanguageListCommand_initValidators = function _SpoWebInstalledLanguageListCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoWebInstalledLanguageListCommand();
//# sourceMappingURL=web-installedlanguage-list.js.map