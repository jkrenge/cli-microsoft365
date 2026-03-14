var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoAppDeployCommand_instances, _SpoAppDeployCommand_initTelemetry, _SpoAppDeployCommand_initOptions, _SpoAppDeployCommand_initValidators, _SpoAppDeployCommand_initOptionSets;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';
class SpoAppDeployCommand extends SpoAppBaseCommand {
    get name() {
        return commands.APP_DEPLOY;
    }
    get description() {
        return 'Deploys the specified app in the specified app catalog';
    }
    constructor() {
        super();
        _SpoAppDeployCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoAppDeployCommand_instances, "m", _SpoAppDeployCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoAppDeployCommand_instances, "m", _SpoAppDeployCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoAppDeployCommand_instances, "m", _SpoAppDeployCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoAppDeployCommand_instances, "m", _SpoAppDeployCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const scope = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const appCatalogUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);
            let res;
            if (args.options.id) {
                if (this.verbose) {
                    await logger.logToStderr(`Using the specified app id ${args.options.id}`);
                }
                res = { UniqueId: args.options.id };
            }
            else {
                if (this.verbose) {
                    await logger.logToStderr(`Looking up app id for app named ${args.options.name}...`);
                }
                const requestOptions = {
                    url: `${appCatalogUrl}/_api/web/GetFolderByServerRelativePath(DecodedUrl='AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                res = await request.get(requestOptions);
            }
            const appId = res.UniqueId;
            if (this.verbose) {
                await logger.logToStderr(`Deploying app...`);
            }
            const requestOptions = {
                url: `${appCatalogUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${appId}')/deploy`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'content-type': 'application/json;odata=nometadata;charset=utf-8'
                },
                data: { 'skipFeatureDeployment': args.options.skipFeatureDeployment || false },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoAppDeployCommand_instances = new WeakSet(), _SpoAppDeployCommand_initTelemetry = function _SpoAppDeployCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: (!(!args.options.id)).toString(),
            name: (!(!args.options.name)).toString(),
            appCatalogUrl: (!(!args.options.appCatalogUrl)).toString(),
            skipFeatureDeployment: args.options.skipFeatureDeployment || false,
            appCatalogScope: (!(!args.options.appCatalogScope)).toString()
        });
    });
}, _SpoAppDeployCommand_initOptions = function _SpoAppDeployCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-n, --name [name]'
    }, {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
    }, {
        option: '--skipFeatureDeployment'
    }, {
        option: '-s, --appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
    });
}, _SpoAppDeployCommand_initValidators = function _SpoAppDeployCommand_initValidators() {
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
}, _SpoAppDeployCommand_initOptionSets = function _SpoAppDeployCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'name'] });
};
export default new SpoAppDeployCommand();
//# sourceMappingURL=app-deploy.js.map