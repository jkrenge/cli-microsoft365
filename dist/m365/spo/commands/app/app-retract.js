var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppRetractCommand_instances, _SpoAppRetractCommand_initTelemetry, _SpoAppRetractCommand_initOptions, _SpoAppRetractCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';
class SpoAppRetractCommand extends SpoAppBaseCommand {
    get name() {
        return commands.APP_RETRACT;
    }
    get description() {
        return 'Retracts the specified app from the specified app catalog';
    }
    constructor() {
        super();
        _SpoAppRetractCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppRetractCommand_instances, "m", _SpoAppRetractCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppRetractCommand_instances, "m", _SpoAppRetractCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppRetractCommand_instances, "m", _SpoAppRetractCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        const scope = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
        const retractApp = async () => {
            try {
                const spoUrl = await spo.getSpoUrl(logger, this.debug);
                const appCatalogSiteUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);
                if (this.verbose) {
                    await logger.logToStderr(`Retracting app...`);
                }
                const requestOptions = {
                    url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(args.options.id)}')/retract`,
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
            await retractApp();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to retract the app ${args.options.id} from the app catalog?` });
            if (result) {
                await retractApp();
            }
        }
    }
}
_SpoAppRetractCommand_instances = new WeakSet(), _SpoAppRetractCommand_initTelemetry = function _SpoAppRetractCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            appCatalogUrl: (!(!args.options.appCatalogUrl)).toString(),
            force: (!(!args.options.force)).toString(),
            appCatalogScope: args.options.appCatalogScope || 'tenant'
        });
    });
}, _SpoAppRetractCommand_initOptions = function _SpoAppRetractCommand_initOptions() {
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
}, _SpoAppRetractCommand_initValidators = function _SpoAppRetractCommand_initValidators() {
    this.validators.push(async (args) => {
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
export default new SpoAppRetractCommand();
//# sourceMappingURL=app-retract.js.map