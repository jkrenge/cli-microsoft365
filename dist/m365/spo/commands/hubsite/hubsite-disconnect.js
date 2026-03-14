var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteDisconnectCommand_instances, _SpoHubSiteDisconnectCommand_initTelemetry, _SpoHubSiteDisconnectCommand_initOptions, _SpoHubSiteDisconnectCommand_initOptionSets, _SpoHubSiteDisconnectCommand_initValidators;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHubSiteDisconnectCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_DISCONNECT;
    }
    get description() {
        return 'Disconnect a hub site from its parent hub site';
    }
    constructor() {
        super();
        _SpoHubSiteDisconnectCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoHubSiteDisconnectCommand_instances, "m", _SpoHubSiteDisconnectCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteDisconnectCommand_instances, "m", _SpoHubSiteDisconnectCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteDisconnectCommand_instances, "m", _SpoHubSiteDisconnectCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteDisconnectCommand_instances, "m", _SpoHubSiteDisconnectCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const disconnectHubSite = async () => {
            try {
                if (this.verbose) {
                    await logger.logToStderr(`Disconnecting hub site '${args.options.id || args.options.title || args.options.url}' from its parent hub site...`);
                }
                const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
                const hubSite = await this.getHubSite(spoAdminUrl, args.options);
                const requestOptions = {
                    url: `${spoAdminUrl}/_api/HubSites/GetById('${hubSite.ID}')`,
                    headers: {
                        accept: 'application/json;odata=nometadata',
                        'if-match': hubSite['odata.etag']
                    },
                    responseType: 'json',
                    data: {
                        ParentHubSiteId: '00000000-0000-0000-0000-000000000000'
                    }
                };
                await request.patch(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await disconnectHubSite();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want disconnect hub site '${args.options.id || args.options.title || args.options.url}' from its parent hub site?` });
            if (result) {
                await disconnectHubSite();
            }
        }
    }
    async getHubSite(spoAdminUrl, options) {
        const requestOptions = {
            headers: {
                accept: 'application/json;odata=minimalmetadata'
            },
            responseType: 'json'
        };
        if (options.id) {
            requestOptions.url = `${spoAdminUrl}/_api/HubSites/GetById('${options.id}')?$select=ID`;
            const result = await request.get(requestOptions);
            return result;
        }
        requestOptions.url = `${spoAdminUrl}/_api/HubSites?$select=ID,Title,SiteUrl&$top=5000`;
        const hubSitesResponse = await request.get(requestOptions);
        const hubSites = hubSitesResponse.value;
        let filteredHubSites = [];
        if (options.title) {
            filteredHubSites = hubSites.filter(site => site.Title.toLowerCase() === options.title.toLowerCase());
        }
        else if (options.url) {
            filteredHubSites = hubSites.filter(site => site.SiteUrl.toLowerCase() === options.url.toLowerCase());
        }
        if (filteredHubSites.length === 0) {
            throw `The specified hub site '${options.title || options.url}' does not exist.`;
        }
        if (filteredHubSites.length > 1) {
            const resultAsKeyValuePair = formatting.convertArrayToHashTable('ID', filteredHubSites);
            const result = await cli.handleMultipleResultsFound(`Multiple hub sites with name '${options.title}' found.`, resultAsKeyValuePair);
            return result;
        }
        return filteredHubSites[0];
    }
}
_SpoHubSiteDisconnectCommand_instances = new WeakSet(), _SpoHubSiteDisconnectCommand_initTelemetry = function _SpoHubSiteDisconnectCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            id: typeof args.options.id !== 'undefined',
            title: typeof args.options.title !== 'undefined',
            url: typeof args.options.url !== 'undefined'
        });
    });
}, _SpoHubSiteDisconnectCommand_initOptions = function _SpoHubSiteDisconnectCommand_initOptions() {
    this.options.unshift({
        option: '-i, --id [id]'
    }, {
        option: '-t, --title [title]'
    }, {
        option: '-u, --url [url]'
    }, {
        option: '-f, --force'
    });
}, _SpoHubSiteDisconnectCommand_initOptionSets = function _SpoHubSiteDisconnectCommand_initOptionSets() {
    this.optionSets.push({ options: ['id', 'title', 'url'] });
}, _SpoHubSiteDisconnectCommand_initValidators = function _SpoHubSiteDisconnectCommand_initValidators() {
    this.validators.push(async (args) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
            return `'${args.options.id}' is not a valid GUID`;
        }
        if (args.options.url) {
            const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.url);
            if (isValidSharePointUrl !== true) {
                return isValidSharePointUrl;
            }
        }
        return true;
    });
};
export default new SpoHubSiteDisconnectCommand();
//# sourceMappingURL=hubsite-disconnect.js.map