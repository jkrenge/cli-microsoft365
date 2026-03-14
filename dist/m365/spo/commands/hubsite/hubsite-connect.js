var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteConnectCommand_instances, _SpoHubSiteConnectCommand_initTelemetry, _SpoHubSiteConnectCommand_initOptions, _SpoHubSiteConnectCommand_initOptionSets, _SpoHubSiteConnectCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';
class SpoHubSiteConnectCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_CONNECT;
    }
    get description() {
        return 'Connect a hub site to a parent hub site';
    }
    constructor() {
        super();
        _SpoHubSiteConnectCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHubSiteConnectCommand_instances, "m", _SpoHubSiteConnectCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteConnectCommand_instances, "m", _SpoHubSiteConnectCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteConnectCommand_instances, "m", _SpoHubSiteConnectCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteConnectCommand_instances, "m", _SpoHubSiteConnectCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Connecting hub site '${args.options.id || args.options.title || args.options.url}' to hub site '${args.options.parentId || args.options.parentTitle || args.options.parentUrl}'...`);
        }
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const hubSites = await this.getHubSites(spoAdminUrl);
            const hubSite = await this.getSpecificHubSite(hubSites, args.options.id, args.options.title, args.options.url);
            const parentHubSite = await this.getSpecificHubSite(hubSites, args.options.parentId, args.options.parentTitle, args.options.parentUrl);
            const requestOptions = {
                url: `${spoAdminUrl}/_api/HubSites/GetById('${hubSite.ID}')`,
                headers: {
                    accept: 'application/json;odata=nometadata',
                    'if-match': hubSite['odata.etag']
                },
                responseType: 'json',
                data: {
                    ParentHubSiteId: parentHubSite.ID
                }
            };
            await request.patch(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getHubSites(spoAdminUrl) {
        const requestOptions = {
            url: `${spoAdminUrl}/_api/HubSites?$select=ID,Title,SiteUrl&$top=5000`,
            headers: {
                accept: 'application/json;odata=minimalmetadata'
            },
            responseType: 'json'
        };
        const hubSites = await request.get(requestOptions);
        return hubSites.value;
    }
    async getSpecificHubSite(hubSites, id, title, url) {
        let filteredHubSites = [];
        if (id) {
            filteredHubSites = hubSites.filter(site => site.ID.toLowerCase() === id.toLowerCase());
        }
        else if (title) {
            filteredHubSites = hubSites.filter(site => site.Title.toLowerCase() === title.toLowerCase());
        }
        else if (url) {
            filteredHubSites = hubSites.filter(site => site.SiteUrl.toLowerCase() === url.toLowerCase());
        }
        if (filteredHubSites.length === 0) {
            throw `The specified hub site '${id || title || url}' does not exist.`;
        }
        if (filteredHubSites.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('ID', filteredHubSites);
            return await cli.handleMultipleResultsFound(`Multiple hub sites with name '${title}' found.`, resultAsKeyValuePair);
        }
        return filteredHubSites[0];
    }
}
_SpoHubSiteConnectCommand_instances = new WeakSet(), _SpoHubSiteConnectCommand_initTelemetry = function _SpoHubSiteConnectCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            url: typeof args.options.url !== 'undefined',
            parentId: typeof args.options.parentId !== 'undefined',
            parentTitle: typeof args.options.parentTitle !== 'undefined',
            parentUrl: typeof args.options.parentUrl !== 'undefined'
        });
    });
}, _SpoHubSiteConnectCommand_initOptions = function _SpoHubSiteConnectCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-u, --url [url]'
    }, {
        option: '--parentId [parentId]'
    }, {
        option: '--parentTitle [parentTitle]'
    }, {
        option: '--parentUrl [parentUrl]'
    });
}, _SpoHubSiteConnectCommand_initOptionSets = function _SpoHubSiteConnectCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'url'] }, { options: ['parentId', 'parentTitle', 'parentUrl'] });
}, _SpoHubSiteConnectCommand_initValidators = function _SpoHubSiteConnectCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID`;
        }
        if (args.options.parentId && !validation.isValidGuid(args.options.parentId)) {
            return `'${args.options.parentId}' is not a valid GUID`;
        }
        if (args.options.url) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.url);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        if (args.options.parentUrl) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.parentUrl);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        return true;
    });
};
export default new SpoHubSiteConnectCommand();
//# sourceMappingURL=hubsite-connect.js.map