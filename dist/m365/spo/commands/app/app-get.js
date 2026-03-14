var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppGetCommand_instances, _SpoAppGetCommand_initTelemetry, _SpoAppGetCommand_initOptions, _SpoAppGetCommand_initValidators, _SpoAppGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';
class SpoAppGetCommand extends SpoAppBaseCommand {
    get name() {
        return commands.APP_GET;
    }
    get description() {
        return 'Gets information about the specific app from the specified app catalog';
    }
    constructor() {
        super();
        _SpoAppGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppGetCommand_instances, "m", _SpoAppGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppGetCommand_instances, "m", _SpoAppGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppGetCommand_instances, "m", _SpoAppGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoAppGetCommand_instances, "m", _SpoAppGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const scope = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const appCatalogSiteUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);
            let appId;
            if (args.options.id) {
                appId = args.options.id;
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr(`Looking up app id for app named ${args.options.name}...`);
                }
                const requestOptions = {
                    url: `${appCatalogSiteUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const app = await request.get(requestOptions);
                appId = app.UniqueId;
            }
            if (this.verbose) {
                await logger.logToStderr(`Retrieving information for app ${appId}...`);
            }
            const requestOptions = {
                url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(appId)}')`,
                headers: {
                    accept: 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const res = await request.get(requestOptions);
            await logger.log(res);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoAppGetCommand_instances = new WeakSet(), _SpoAppGetCommand_initTelemetry = function _SpoAppGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: (!(!args.options.id)).toString(),
            name: (!(!args.options.name)).toString(),
            appCatalogUrl: (!(!args.options.appCatalogUrl)).toString(),
            appCatalogScope: args.options.appCatalogScope || 'tenant'
        });
    });
}, _SpoAppGetCommand_initOptions = function _SpoAppGetCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
    }, {
        option: '-s, --appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
    });
}, _SpoAppGetCommand_initValidators = function _SpoAppGetCommand_initValidators() {
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
        }
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.appCatalogUrl) {
            return validation.isValidSharePointUrl(args.options.appCatalogUrl);
        }
        return true;
    });
}, _SpoAppGetCommand_initOptionSets = function _SpoAppGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new SpoAppGetCommand();
//# sourceMappingURL=app-get.js.map