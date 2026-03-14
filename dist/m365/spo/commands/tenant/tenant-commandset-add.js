var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoTenantCommandSetAddCommand_instances, _a, _SpoTenantCommandSetAddCommand_initTelemetry, _SpoTenantCommandSetAddCommand_initOptions, _SpoTenantCommandSetAddCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoListItemAddCommand from '../listitem/listitem-add.js';
import spoListItemListCommand from '../listitem/listitem-list.js';
import spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get.js';
class SpoTenantCommandSetAddCommand extends SpoCommand {
    get name() {
        return commands.TENANT_COMMANDSET_ADD;
    }
    get description() {
        return 'Add a ListView Command Set as a tenant-wide extension.';
    }
    constructor() {
        super();
        _SpoTenantCommandSetAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetAddCommand_instances, "m", _SpoTenantCommandSetAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetAddCommand_instances, "m", _SpoTenantCommandSetAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoTenantCommandSetAddCommand_instances, "m", _SpoTenantCommandSetAddCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        try {
            const appCatalogUrl = await this.getAppCatalogUrl(logger);
            const componentManifest = await this.getComponentManifest(appCatalogUrl, args.options.clientSideComponentId, logger);
            const clientComponentManifest = JSON.parse(componentManifest.ClientComponentManifest);
            if (clientComponentManifest.extensionType !== "ListViewCommandSet") {
                throw `The extension type of this component is not of type 'ListViewCommandSet' but of type '${clientComponentManifest.extensionType}'`;
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
            throw 'Cannot add tenant-wide ListView Command Set as app catalog cannot be found';
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
            TenantWideExtensionLocation: this.getLocation(options.location),
            TenantWideExtensionSequence: 0,
            TenantWideExtensionListTemplate: this.getListTemplate(options.listType),
            TenantWideExtensionComponentProperties: options.clientSideComponentProperties || '',
            TenantWideExtensionWebTemplate: options.webTemplate || '',
            TenantWideExtensionDisabled: false,
            verbose: this.verbose,
            debug: this.debug,
            output: options.output
        };
        await cli.executeCommand(spoListItemAddCommand, { options: { ...commandOptions, _: [] } });
    }
    getLocation(location) {
        switch (location) {
            case 'Both':
                return 'ClientSideExtension.ListViewCommandSet';
            case 'ContextMenu':
                return 'ClientSideExtension.ListViewCommandSet.ContextMenu';
            default:
                return 'ClientSideExtension.ListViewCommandSet.CommandBar';
        }
    }
    getListTemplate(listTemplate) {
        switch (listTemplate) {
            case 'SitePages':
                return '119';
            case 'Library':
                return '101';
            default:
                return '100';
        }
    }
}
_a = SpoTenantCommandSetAddCommand, _SpoTenantCommandSetAddCommand_instances = new WeakSet(), _SpoTenantCommandSetAddCommand_initTelemetry = function _SpoTenantCommandSetAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listType: args.options.listType,
            clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
            webTemplate: typeof args.options.webTemplate !== 'undefined',
            location: args.options.location
        });
    });
}, _SpoTenantCommandSetAddCommand_initOptions = function _SpoTenantCommandSetAddCommand_initOptions() {
    this.options.unshift({
        option: '-t, --title <title>'
    }, {
        option: '-l, --listType <listType>',
        autocomplete: _a.listTypes
    }, {
        option: '-i, --clientSideComponentId <clientSideComponentId>'
    }, {
        option: '-p, --clientSideComponentProperties [clientSideComponentProperties]'
    }, {
        option: '-w, --webTemplate [webTemplate]'
    }, {
        option: '--location [location]',
        autocomplete: _a.locations
    });
}, _SpoTenantCommandSetAddCommand_initValidators = function _SpoTenantCommandSetAddCommand_initValidators() {
    this.validators.push(async (args) => {
        if (!validation.isValidGuid(args.options.clientSideComponentId)) {
            return `${args.options.clientSideComponentId} is not a valid GUID`;
        }
        if (_a.listTypes.indexOf(args.options.listType) < 0) {
            return `${args.options.listType} is not a valid list type. Allowed values are ${_a.listTypes.join(', ')}`;
        }
        if (args.options.location && _a.locations.indexOf(args.options.location) < 0) {
            return `${args.options.location} is not a valid location. Allowed values are ${_a.locations.join(', ')}`;
        }
        return true;
    });
};
SpoTenantCommandSetAddCommand.listTypes = ['List', 'Library', 'SitePages'];
SpoTenantCommandSetAddCommand.locations = ['ContextMenu', 'CommandBar', 'Both'];
export default new SpoTenantCommandSetAddCommand();
//# sourceMappingURL=tenant-commandset-add.js.map