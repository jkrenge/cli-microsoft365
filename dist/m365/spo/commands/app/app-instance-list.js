var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppInStanceListCommand_instances, _SpoAppInStanceListCommand_initOptions, _SpoAppInStanceListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';
class SpoAppInStanceListCommand extends SpoAppBaseCommand {
    get name() {
        return commands.APP_INSTANCE_LIST;
    }
    get description() {
        return 'Retrieve apps installed in a site';
    }
    defaultProperties() {
        return [`Title`, `AppId`];
    }
    constructor() {
        super();
        _SpoAppInStanceListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppInStanceListCommand_instances, "m", _SpoAppInStanceListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppInStanceListCommand_instances, "m", _SpoAppInStanceListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving installed apps in site at ${args.options.siteUrl}...`);
        }
        try {
            const apps = await odata.getAllItems(`${args.options.siteUrl}/_api/web/AppTiles?$filter=AppType eq 3`);
            await logger.log(apps);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoAppInStanceListCommand_instances = new WeakSet(), _SpoAppInStanceListCommand_initOptions = function _SpoAppInStanceListCommand_initOptions() {
    this.options.unshift({
        option: '-u, --siteUrl <siteUrl>'
    });
}, _SpoAppInStanceListCommand_initValidators = function _SpoAppInStanceListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        return true;
    });
};
export default new SpoAppInStanceListCommand();
//# sourceMappingURL=app-instance-list.js.map