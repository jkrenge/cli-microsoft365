var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppUninstallCommand_instances, _SpoAppUninstallCommand_initTelemetry, _SpoAppUninstallCommand_initOptions, _SpoAppUninstallCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoAppUninstallCommand extends SpoCommand {
    get name() {
        return commands.APP_UNINSTALL;
    }
    get description() {
        return 'Uninstalls an app from the site';
    }
    constructor() {
        super();
        _SpoAppUninstallCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppUninstallCommand_instances, "m", _SpoAppUninstallCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppUninstallCommand_instances, "m", _SpoAppUninstallCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppUninstallCommand_instances, "m", _SpoAppUninstallCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const uninstallApp = async () => {
            const scope = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
            if (this.verbose) {
                await logger.logToStderr(`Uninstalling app '${args.options.id}' from the site '${args.options.siteUrl}'...`);
            }
            const requestOptions = {
                url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(args.options.id)}')/uninstall`,
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
        };
        if (args.options.force) {
            await uninstallApp();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to uninstall the app ${args.options.id} from site ${args.options.siteUrl}?` });
            if (result) {
                await uninstallApp();
            }
        }
    }
}
_SpoAppUninstallCommand_instances = new WeakSet(), _SpoAppUninstallCommand_initTelemetry = function _SpoAppUninstallCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            force: (!(!args.options.force)).toString(),
            appCatalogScope: args.options.appCatalogScope || 'tenant'
        });
    });
}, _SpoAppUninstallCommand_initOptions = function _SpoAppUninstallCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-s, --siteUrl <siteUrl>'
    }, {
        option: '--appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
    }, {
        option: '-f, --force'
    });
}, _SpoAppUninstallCommand_initValidators = function _SpoAppUninstallCommand_initValidators() {
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
export default new SpoAppUninstallCommand();
//# sourceMappingURL=app-uninstall.js.map