var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppRemoveCommand_instances, _SpoAppRemoveCommand_initTelemetry, _SpoAppRemoveCommand_initOptions, _SpoAppRemoveCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';
class SpoAppRemoveCommand extends SpoAppBaseCommand {
    get name() {
        return commands.APP_REMOVE;
    }
    get description() {
        return 'Removes the specified app from the specified app catalog';
    }
    constructor() {
        super();
        _SpoAppRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppRemoveCommand_instances, "m", _SpoAppRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppRemoveCommand_instances, "m", _SpoAppRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppRemoveCommand_instances, "m", _SpoAppRemoveCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const scope = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
        const removeApp = async () => {
            try {
                const spoUrl = await spo.getSpoUrl(logger, this.debug);
                const appCatalogUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);
                if (this.debug) {
                    await logger.logToStderr(`Retrieved app catalog URL ${appCatalogUrl}. Removing app from the app catalog...`);
                }
                const requestOptions = {
                    url: `${appCatalogUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(args.options.id)}')/remove`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    }
                };
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataPromise(err);
            }
        };
        if (args.options.force) {
            await removeApp();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the app ${args.options.id} from the app catalog?` });
            if (result) {
                await removeApp();
            }
        }
    }
}
_SpoAppRemoveCommand_instances = new WeakSet(), _SpoAppRemoveCommand_initTelemetry = function _SpoAppRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appCatalogUrl: (!(!args.options.appCatalogUrl)).toString(),
            force: (!(!args.options.force)).toString(),
            appCatalogScope: args.options.appCatalogScope || 'tenant'
        });
    });
}, _SpoAppRemoveCommand_initOptions = function _SpoAppRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id <id>'
    }, {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
    }, {
        option: '-s, --appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
    }, {
        option: '-f, --force'
    });
}, _SpoAppRemoveCommand_initValidators = function _SpoAppRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        // verify either 'tenant' or 'sitecollection' specified if scope provided
        if (args.options.appCatalogScope) {
            const testScope = args.options.appCatalogScope.toLowerCase();
            if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
                return `appCatalogScope must be either 'tenant' or 'sitecollection' if specified`;
            }
            if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
                return `You must specify appCatalogUrl when the appCatalogScope is sitecollection`;
            }
        }
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.appCatalogUrl) {
            return validation.isValidSharePointUrl(args.options.appCatalogUrl);
        }
        return true;
    });
};
export default new SpoAppRemoveCommand();
//# sourceMappingURL=app-remove.js.map