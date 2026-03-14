var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppInstallCommand_instances, _SpoAppInstallCommand_initTelemetry, _SpoAppInstallCommand_initOptions, _SpoAppInstallCommand_initValidators;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoAppInstallCommand extends SpoCommand {
    get name() {
        return commands.APP_INSTALL;
    }
    get description() {
        return 'Installs an app from the specified app catalog in the site';
    }
    constructor() {
        super();
        _SpoAppInstallCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppInstallCommand_instances, "m", _SpoAppInstallCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppInstallCommand_instances, "m", _SpoAppInstallCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppInstallCommand_instances, "m", _SpoAppInstallCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const scope = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
        if (this.verbose) {
            await logger.logToStderr(`Installing app '${args.options.id}' in site '${args.options.siteUrl}'...`);
        }
        const requestOptions = {
            url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(args.options.id)}')/install`,
            headers: {
                accept: 'application/json;odata=nometadata'
            }
        };
        try {
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataPromise(err);
        }
    }
}
_SpoAppInstallCommand_instances = new WeakSet(), _SpoAppInstallCommand_initTelemetry = function _SpoAppInstallCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appCatalogScope: args.options.appCatalogScope || 'tenant'
        });
    });
}, _SpoAppInstallCommand_initOptions = function _SpoAppInstallCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-s, --siteUrl <siteUrl>'
    }, {
        option: '--appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
    });
}, _SpoAppInstallCommand_initValidators = function _SpoAppInstallCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.appCatalogScope) {
            const testScope = args.options.appCatalogScope.toLowerCase();
            if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
                return `appCatalogScope must be either 'tenant' or 'sitecollection' if specified`;
            }
        }
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        return validation.isValidSharePointUrl(args.options.siteUrl);
    });
};
export default new SpoAppInstallCommand();
//# sourceMappingURL=app-install.js.map