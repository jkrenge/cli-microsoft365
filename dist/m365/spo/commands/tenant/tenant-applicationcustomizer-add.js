var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantApplicationCustomizerAddCommand_instances, _SpoTenantApplicationCustomizerAddCommand_initTelemetry, _SpoTenantApplicationCustomizerAddCommand_initOptions, _SpoTenantApplicationCustomizerAddCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoListItemAddCommand from '../listitem/listitem-add.js';
import spoListItemListCommand from '../listitem/listitem-list.js';
import spoTenantAppCatalogUrlGetCommand from '../tenant/tenant-appcatalogurl-get.js';
class SpoTenantApplicationCustomizerAddCommand extends SpoCommand {
    get name() {
        return commands.TENANT_APPLICATIONCUSTOMIZER_ADD;
    }
    get description() {
        return 'Add an application customizer as a tenant wide extension.';
    }
    constructor() {
        super();
        _SpoTenantApplicationCustomizerAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerAddCommand_instances, "m", _SpoTenantApplicationCustomizerAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerAddCommand_instances, "m", _SpoTenantApplicationCustomizerAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantApplicationCustomizerAddCommand_instances, "m", _SpoTenantApplicationCustomizerAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appCatalogUrl = await this.getAppCatalogUrl(logger);
            const componentManifest = await this.getComponentManifest(appCatalogUrl, args.options.clientSideComponentId, logger);
            const clientComponentManifest = JSON.parse(componentManifest.ClientComponentManifest);
            if (clientComponentManifest.extensionType !== "ApplicationCustomizer") {
                throw `The extension type of this component is not of type 'ApplicationCustomizer' but of type '${clientComponentManifest.extensionType}'`;
            }
            const solution = await this.getSolutionFromAppCatalog(appCatalogUrl, componentManifest.SolutionId, logger);
            if (!solution.ContainsTenantWideExtension) {
                throw `The solution does not contain an extension that can be deployed to all sites. Make sure that you've entered the correct component Id.`;
            }
            else if (!solution.SkipFeatureDeployment) {
                throw 'The solution has not been deployed to all sites. Make sure to deploy this solution to all sites.';
            }
            await this.addTenantWideExtension(appCatalogUrl, args.options, logger);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAppCatalogUrl(logger) {
        const spoTenantAppCatalogUrlGetCommandOutput = await cli.executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand, { options: { output: 'text', _: [] } });
        if (this.verbose) {
            await logger.logToStderr(spoTenantAppCatalogUrlGetCommandOutput.stderr);
        }
        const appCatalogUrl = spoTenantAppCatalogUrlGetCommandOutput.stdout;
        if (!appCatalogUrl) {
            throw 'Cannot add tenant-wide application customizer as app catalog cannot be found';
        }
        if (this.verbose) {
            await logger.logToStderr(`Got tenant app catalog url: ${appCatalogUrl}`);
        }
        return appCatalogUrl;
    }
    async getComponentManifest(appCatalogUrl, clientSideComponentId, logger) {
        if (this.verbose) {
            await logger.logToStderr('Retrieving component manifest item from the ComponentManifests list on the app catalog site so that we get the solution id');
        }
        const camlQuery = `<View><ViewFields><FieldRef Name='ClientComponentId'></FieldRef><FieldRef Name='SolutionId'></FieldRef><FieldRef Name='ClientComponentManifest'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='ClientComponentId' /><Value Type='Guid'>${clientSideComponentId}</Value></Eq></Where></Query></View>`;
        const commandOptions = {
            webUrl: appCatalogUrl,
            listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`,
            camlQuery: camlQuery,
            verbose: this.verbose,
            debug: this.debug,
            output: 'json'
        };
        const output = await cli.executeCommandWithOutput(spoListItemListCommand, { options: { ...commandOptions, _: [] } });
        if (this.verbose) {
            await logger.logToStderr(output.stderr);
        }
        const outputParsed = JSON.parse(output.stdout);
        if (outputParsed.length === 0) {
            throw 'No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog';
        }
        return outputParsed[0];
    }
    async getSolutionFromAppCatalog(appCatalogUrl, solutionId, logger) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving solution with id ${solutionId} from the application catalog`);
        }
        const camlQuery = `<View><ViewFields><FieldRef Name='SkipFeatureDeployment'></FieldRef><FieldRef Name='ContainsTenantWideExtension'></FieldRef></ViewFields><Query><Where><Eq><FieldRef Name='AppProductID' /><Value Type='Guid'>${solutionId}</Value></Eq></Where></Query></View>`;
        const commandOptions = {
            webUrl: appCatalogUrl,
            listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`,
            camlQuery: camlQuery,
            verbose: this.verbose,
            debug: this.debug,
            output: 'json'
        };
        const output = await cli.executeCommandWithOutput(spoListItemListCommand, { options: { ...commandOptions, _: [] } });
        if (this.verbose) {
            await logger.logToStderr(output.stderr);
        }
        const outputParsed = JSON.parse(output.stdout);
        if (outputParsed.length === 0) {
            throw `No component found with the solution id ${solutionId}. Make sure that the solution is available in the app catalog`;
        }
        return outputParsed[0];
    }
    async addTenantWideExtension(appCatalogUrl, options, logger) {
        if (this.verbose) {
            await logger.logToStderr('Pre-checks finished. Adding tenant wide extension to the TenantWideExtensions list');
        }
        const commandOptions = {
            webUrl: appCatalogUrl,
            listUrl: `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/TenantWideExtensions`,
            Title: options.title,
            TenantWideExtensionComponentId: options.clientSideComponentId,
            TenantWideExtensionLocation: 'ClientSideExtension.ApplicationCustomizer',
            TenantWideExtensionSequence: 0,
            TenantWideExtensionListTemplate: 0,
            TenantWideExtensionComponentProperties: options.clientSideComponentProperties || '',
            TenantWideExtensionWebTemplate: options.webTemplate || '',
            TenantWideExtensionDisabled: false,
            verbose: this.verbose,
            debug: this.debug,
            output: options.output
        };
        await cli.executeCommand(spoListItemAddCommand, { options: { ...commandOptions, _: [] } });
    }
}
_SpoTenantApplicationCustomizerAddCommand_instances = new WeakSet(), _SpoTenantApplicationCustomizerAddCommand_initTelemetry = function _SpoTenantApplicationCustomizerAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
            webTemplate: typeof args.options.webTemplate !== 'undefined'
        });
    });
}, _SpoTenantApplicationCustomizerAddCommand_initOptions = function _SpoTenantApplicationCustomizerAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
    }, {
        option: '-i, --clientSideComponentId <clientSideComponentId>'
    }, {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
    }, {
        option: '-w, --webTemplate [webTemplate]'
    });
}, _SpoTenantApplicationCustomizerAddCommand_initValidators = function _SpoTenantApplicationCustomizerAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        return true;
    });
};
export default new SpoTenantApplicationCustomizerAddCommand();
//# sourceMappingURL=tenant-applicationcustomizer-add.js.map