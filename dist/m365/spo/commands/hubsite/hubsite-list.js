var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoHubSiteListCommand_instances, _SpoHubSiteListCommand_initTelemetry, _SpoHubSiteListCommand_initOptions;
import request from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoHubSiteListCommand extends SpoCommand {
    get name() {
        return commands.HUBSITE_LIST;
    }
    get description() {
        return 'Lists hub sites in the current tenant';
    }
    defaultProperties() {
        return ['ID', 'SiteUrl', 'Title'];
    }
    constructor() {
        super();
        _SpoHubSiteListCommand_instances.add(this);
        this.batchSize = 30;
        this.batchSize = 30;
        __classPrivateFieldGet(this, _SpoHubSiteListCommand_instances, "m", _SpoHubSiteListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoHubSiteListCommand_instances, "m", _SpoHubSiteListCommand_initOptions).call(this);
    }
    async commandAction(logger, args) {
        try {
            const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const hubSitesResult = await odata.getAllItems(`${spoAdminUrl}/_api/hubsites`);
            const hubSites = hubSitesResult;
            if (!(args.options.withAssociatedSites !== true || args.options.output && args.options.output !== 'json')) {
                if (this.debug) {
                    await logger.logToStderr('Retrieving associated sites...');
                    await logger.logToStderr('');
                }
                const requestOptions = {
                    url: `${spoAdminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/RenderListDataAsStream`,
                    headers: {
                        accept: 'application/json;odata=nometadata'
                    },
                    responseType: 'json',
                    data: {
                        parameters: {
                            ViewXml: "<View><Query><Where><And><And><IsNull><FieldRef Name=\"TimeDeleted\"/></IsNull><Neq><FieldRef Name=\"State\"/><Value Type='Integer'>0</Value></Neq></And><Neq><FieldRef Name=\"HubSiteId\"/><Value Type='Text'>{00000000-0000-0000-0000-000000000000}</Value></Neq></And></Where><OrderBy><FieldRef Name='Title' Ascending='true' /></OrderBy></Query><ViewFields><FieldRef Name=\"Title\"/><FieldRef Name=\"SiteUrl\"/><FieldRef Name=\"SiteId\"/><FieldRef Name=\"HubSiteId\"/></ViewFields><RowLimit Paged=\"TRUE\">" + this.batchSize + "</RowLimit></View>",
                            DatesInUtc: true
                        }
                    }
                };
                if (this.debug) {
                    await logger.logToStderr(`Will retrieve associated sites (including the hub sites) in batches of ${this.batchSize}`);
                }
                const res = await this.getSites(requestOptions, requestOptions.url, logger);
                if (res) {
                    hubSites.forEach(h => {
                        const filteredSites = res.filter(f => {
                            // Only include sites of which the Site Id is not the same as the
                            // Hub Site ID (as this site is the actual hub site) and of which the
                            // Hub Site ID matches the ID of the Hub
                            return f.SiteId !== f.HubSiteId
                                && f.HubSiteId.toUpperCase() === `{${h.ID.toUpperCase()}}`;
                        });
                        h.AssociatedSites = filteredSites.map(a => {
                            return {
                                Title: a.Title,
                                SiteUrl: a.SiteUrl
                            };
                        });
                    });
                }
            }
            await logger.log(hubSites);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getSites(reqOptions, nonPagedUrl, logger, sites = [], batchNumber = 0) {
        const res = await request.post(reqOptions);
        batchNumber++;
        const retrievedSites = res.Row.length > 0 ? sites.concat(res.Row) : sites;
        if (this.debug) {
            await logger.logToStderr(res);
            await logger.logToStderr(`Retrieved ${res.Row.length} sites in batch ${batchNumber}`);
        }
        if (res.NextHref) {
            reqOptions.url = nonPagedUrl + res.NextHref;
            if (this.debug) {
                await logger.logToStderr(`Url for next batch of sites: ${reqOptions.url}`);
            }
            return this.getSites(reqOptions, nonPagedUrl, logger, retrievedSites, batchNumber);
        }
        else {
            if (this.debug) {
                await logger.logToStderr(`Retrieved ${retrievedSites.length} sites in total`);
            }
            return retrievedSites;
        }
    }
}
_SpoHubSiteListCommand_instances = new WeakSet(), _SpoHubSiteListCommand_initTelemetry = function _SpoHubSiteListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            withAssociatedSites: args.options.withAssociatedSites === true
        });
    });
}, _SpoHubSiteListCommand_initOptions = function _SpoHubSiteListCommand_initOptions() {
    this.options.unshift({
        option: '--withAssociatedSites'
    });
};
export default new SpoHubSiteListCommand();
//# sourceMappingURL=hubsite-list.js.map