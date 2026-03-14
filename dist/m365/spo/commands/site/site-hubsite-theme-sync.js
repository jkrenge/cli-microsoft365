var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteHubSiteThemeSyncCommand_instances, _SpoSiteHubSiteThemeSyncCommand_initOptions, _SpoSiteHubSiteThemeSyncCommand_initValidators;
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoSiteHubSiteThemeSyncCommand extends SpoCommand {
    get name() {
        return commands.SITE_HUBSITE_THEME_SYNC;
    }
    get description() {
        return 'Applies any theme updates from the hub site the site is connected to.';
    }
    constructor() {
        super();
        _SpoSiteHubSiteThemeSyncCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoSiteHubSiteThemeSyncCommand_instances, "m", _SpoSiteHubSiteThemeSyncCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteHubSiteThemeSyncCommand_instances, "m", _SpoSiteHubSiteThemeSyncCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr('Syncing hub site theme...');
            }
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/SyncHubSiteTheme`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoSiteHubSiteThemeSyncCommand_instances = new WeakSet(), _SpoSiteHubSiteThemeSyncCommand_initOptions = function _SpoSiteHubSiteThemeSyncCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    });
}, _SpoSiteHubSiteThemeSyncCommand_initValidators = function _SpoSiteHubSiteThemeSyncCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.webUrl));
};
export default new SpoSiteHubSiteThemeSyncCommand();
//# sourceMappingURL=site-hubsite-theme-sync.js.map