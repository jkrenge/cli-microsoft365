var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteGetCommand_instances, _SpoHubSiteGetCommand_initTelemetry, _SpoHubSiteGetCommand_initOptions, _SpoHubSiteGetCommand_initValidators, _SpoHubSiteGetCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoListItemListCommand from '../listitem/listitem-list.js';
class SpoHubSiteGetCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_GET;
    }
    get description() {
        return 'Gets information about the specified hub site';
    }
    constructor() {
        super();
        _SpoHubSiteGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHubSiteGetCommand_instances, "m", _SpoHubSiteGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteGetCommand_instances, "m", _SpoHubSiteGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteGetCommand_instances, "m", _SpoHubSiteGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteGetCommand_instances, "m", _SpoHubSiteGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoUrl = await spo.getSpoUrl(logger, this.debug);
            const hubSite = args.options.id ? await this.getHubSiteById(spoUrl, args.options) : await this.getHubSite(spoUrl, args.options);
            if (args.options.withAssociatedSites && (args.options.output && args.options.output !== 'json')) {
                throw 'withAssociatedSites option is only allowed with json output mode';
            }
            if (args.options.withAssociatedSites === true && args.options.output && !cli.shouldTrimOutput(args.options.output)) {
                const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
                const associatedSitesCommandOutput = await this.getAssociatedSites(spoAdminUrl, hubSite.SiteId, logger, args);
                const associatedSites = JSON.parse(associatedSitesCommandOutput.stdout);
                hubSite.AssociatedSites = associatedSites.filter(s => s.SiteId !== hubSite.SiteId);
            }
            await logger.log(hubSite);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getAssociatedSites(spoAdminUrl, hubSiteId, logger, args) {
        const options = {
            output: 'json',
            debug: args.options.debug,
            verbose: args.options.verbose,
            listTitle: 'DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS',
            webUrl: spoAdminUrl,
            filter: `HubSiteId eq '${hubSiteId}'`,
            fields: 'Title,SiteUrl,SiteId'
        };
        return cli.executeCommandWithOutput(spoListItemListCommand, { options: { ...options, _: [] } });
    }
    async getHubSiteById(spoUrl, options) {
        const requestOptions = {
            url: `${spoUrl}/_api/hubsites/getbyid('${options.id}')`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.get(requestOptions);
    }
    async getHubSite(spoUrl, options) {
        const requestOptions = {
            url: `${spoUrl}/_api/hubsites`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        let hubSites = response.value;
        if (options.title) {
            hubSites = hubSites.filter(site => site.Title.toLocaleLowerCase() === options.title.toLocaleLowerCase());
        }
        else if (options.url) {
            hubSites = hubSites.filter(site => site.SiteUrl.toLocaleLowerCase() === options.url.toLocaleLowerCase());
        }
        if (hubSites.length === 0) {
            throw `The specified hub site ${options.title || options.url} does not exist`;
        }
        if (hubSites.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('ID', hubSites);
            return await cli.handleMultipleResultsFound(`Multiple hub sites with ${options.title || options.url} found.`, resultAsKeyValuePair);
        }
        return hubSites[0];
    }
}
_SpoHubSiteGetCommand_instances = new WeakSet(), _SpoHubSiteGetCommand_initTelemetry = function _SpoHubSiteGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            url: typeof args.options.url !== 'undefined',
            withAssociatedSites: args.options.withAssociatedSites === true
        });
    });
}, _SpoHubSiteGetCommand_initOptions = function _SpoHubSiteGetCommand_initOptions() {
    this.options.unshift({ option: '-i, --id [id]' }, { option: '-t, --title [title]' }, { option: '-u, --url [url]' }, { option: '--withAssociatedSites' });
}, _SpoHubSiteGetCommand_initValidators = function _SpoHubSiteGetCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
        }
        if (args.options.url) {
            return validation.isValidSharePointUrl(args.options.url);
        }
        return true;
    });
}, _SpoHubSiteGetCommand_initOptionSets = function _SpoHubSiteGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'url'] });
};
export default new SpoHubSiteGetCommand();
//# sourceMappingURL=hubsite-get.js.map