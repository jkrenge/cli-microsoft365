var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppListCommand_instances, _SpoAppListCommand_initTelemetry, _SpoAppListCommand_initOptions, _SpoAppListCommand_initValidators;
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';
class SpoAppListCommand extends SpoAppBaseCommand {
    get name() {
        return commands.APP_LIST;
    }
    get description() {
        return 'Lists apps from the specified app catalog';
    }
    defaultProperties() {
        return [`Title`, `ID`, `Deployed`, `AppCatalogVersion`];
    }
    constructor() {
        super();
        _SpoAppListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppListCommand_instances, "m", _SpoAppListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppListCommand_instances, "m", _SpoAppListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppListCommand_instances, "m", _SpoAppListCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const scope = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const appCatalogSiteUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);
            if (this.verbose) {
                await logger.logToStderr(`Retrieving apps...`);
            }
            const apps = await odata.getAllItems(`${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps`);
            await logger.log(apps);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoAppListCommand_instances = new WeakSet(), _SpoAppListCommand_initTelemetry = function _SpoAppListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appCatalogUrl: (!(!args.options.appCatalogUrl)).toString(),
            appCatalogScope: args.options.appCatalogScope || 'tenant'
        });
    });
}, _SpoAppListCommand_initOptions = function _SpoAppListCommand_initOptions() {
    this.options.unshift({
        option: '-s, --appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
    }, {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
    });
}, _SpoAppListCommand_initValidators = function _SpoAppListCommand_initValidators() {
    this.validators.push(async (args) => {
        // verify either 'tenant' or 'sitecollection' specified if scope provided
        if (args.options.appCatalogScope) {
            const testScope = args.options.appCatalogScope.toLowerCase();
            if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
                return `appCatalogScope must be either 'tenant' or 'sitecollection'`;
            }
            if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
                return `You must specify appCatalogUrl when the appCatalogScope is sitecollection`;
            }
            if (args.options.appCatalogUrl) {
                return validation.isValidSharePointUrl(args.options.appCatalogUrl);
            }
        }
        return true;
    });
};
export default new SpoAppListCommand();
//# sourceMappingURL=app-list.js.map