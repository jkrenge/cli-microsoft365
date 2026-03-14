var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoSiteRemoveCommand_instances, _SpoSiteRemoveCommand_initTelemetry, _SpoSiteRemoveCommand_initOptions, _SpoSiteRemoveCommand_initValidators, _SpoSiteRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { setTimeout } from 'timers/promises';
import { urlUtil } from '../../../../utils/urlUtil.js';
class SpoSiteRemoveCommand extends SpoCommand {
    get name() {
        return commands.SITE_REMOVE;
    }
    get description() {
        return 'Removes the specified site';
    }
    constructor() {
        super();
        _SpoSiteRemoveCommand_instances.add(this);
        this.pollingInterval = 5000;
        __classPrivateFieldGet(this, _SpoSiteRemoveCommand_instances, "m", _SpoSiteRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoSiteRemoveCommand_instances, "m", _SpoSiteRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoSiteRemoveCommand_instances, "m", _SpoSiteRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoSiteRemoveCommand_instances, "m", _SpoSiteRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeSite(logger, args.options);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the site '${args.options.url}'?` });
            if (result) {
                await this.removeSite(logger, args.options);
            }
        }
    }
    async removeSite(logger, options) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Removing site '${options.url}'...`);
            }
            this.spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
            const siteUrl = urlUtil.removeTrailingSlashes(options.url);
            const siteDetails = await this.getSiteDetails(logger, siteUrl);
            const isGroupSite = siteDetails.GroupId && siteDetails.GroupId !== '00000000-0000-0000-0000-000000000000';
            if (options.fromRecycleBin) {
                if (!siteDetails.TimeDeleted) {
                    throw `Site is currently not in the recycle bin. Remove --fromRecycleBin if you want to remove it as active site.`;
                }
                if (isGroupSite) {
                    if (this.verbose) {
                        await logger.logToStderr(`Checking if group '${siteDetails.GroupId}' is already permanently deleted from recycle bin.`);
                    }
                    const isGroupInRecycleBin = await this.isGroupInEntraRecycleBin(logger, siteDetails.GroupId);
                    if (isGroupInRecycleBin) {
                        await this.removeGroupFromEntraRecycleBin(logger, siteDetails.GroupId);
                    }
                }
                await this.deleteSiteFromSharePointRecycleBin(logger, siteUrl);
            }
            else {
                if (siteDetails.TimeDeleted) {
                    throw `Site is already in the recycle bin. Use --fromRecycleBin to permanently delete it.`;
                }
                if (isGroupSite) {
                    await this.deleteGroupifiedSite(logger, siteUrl);
                    if (options.skipRecycleBin) {
                        let isGroupInRecycleBin = await this.isGroupInEntraRecycleBin(logger, siteDetails.GroupId);
                        let amountOfPolls = 0;
                        while (!isGroupInRecycleBin && amountOfPolls < 20) {
                            await setTimeout(this.pollingInterval);
                            isGroupInRecycleBin = await this.isGroupInEntraRecycleBin(logger, siteDetails.GroupId);
                            amountOfPolls++;
                        }
                        if (isGroupInRecycleBin) {
                            await this.removeGroupFromEntraRecycleBin(logger, siteDetails.GroupId);
                        }
                    }
                }
                else {
                    await this.deleteNonGroupSite(logger, siteUrl);
                }
                if (options.skipRecycleBin) {
                    await this.deleteSiteFromSharePointRecycleBin(logger, siteUrl);
                }
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async removeGroupFromEntraRecycleBin(logger, groupId) {
        if (this.verbose) {
            await logger.logToStderr(`Permanently deleting group '${groupId}'.`);
        }
        const requestOptions = {
            url: `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${groupId}`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        return request.delete(requestOptions);
    }
    async isGroupInEntraRecycleBin(logger, groupId) {
        if (this.verbose) {
            await logger.logToStderr(`Checking if group '${groupId}' is in the Microsoft Entra recycle bin.`);
        }
        const requestOptions = {
            url: `https://graph.microsoft.com/v1.0/directory/deletedItems/Microsoft.Graph.Group/${groupId}?$select=id`,
            headers: {
                accept: 'application/json;odata.metadata=none'
            },
            responseType: 'json'
        };
        try {
            await request.get(requestOptions);
            return true;
        }
        catch (err) {
            if (err.response?.status === 404) {
                return false;
            }
            throw err;
        }
    }
    async deleteNonGroupSite(logger, siteUrl) {
        if (this.verbose) {
            await logger.logToStderr(`Deleting site.`);
        }
        const requestOptions = {
            url: `${this.spoAdminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveSite`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            data: {
                siteUrl: siteUrl
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async deleteSiteFromSharePointRecycleBin(logger, url) {
        if (this.verbose) {
            await logger.logToStderr(`Permanently deleting site from the recycle bin.`);
        }
        const requestOptions = {
            url: `${this.spoAdminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveDeletedSite`,
            headers: {
                accept: 'application/json;odata=nometadata',
                'Content-Type': 'application/json'
            },
            data: {
                siteUrl: url
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
    async getSiteDetails(logger, url) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving site info.`);
        }
        const sites = await odata.getAllItems(`${this.spoAdminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/items?$filter=SiteUrl eq '${formatting.encodeQueryParameter(url)}'&$select=GroupId,TimeDeleted,SiteId`);
        if (sites.length === 0) {
            throw `Site not found in the tenant.`;
        }
        return sites[0];
    }
    async deleteGroupifiedSite(logger, siteUrl) {
        if (this.verbose) {
            await logger.logToStderr(`Removing groupified site.`);
        }
        const requestOptions = {
            url: `${this.spoAdminUrl}/_api/GroupSiteManager/Delete?siteUrl='${formatting.encodeQueryParameter(siteUrl)}'`,
            headers: {
                accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return request.post(requestOptions);
    }
}
_SpoSiteRemoveCommand_instances = new WeakSet(), _SpoSiteRemoveCommand_initTelemetry = function _SpoSiteRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            skipRecycleBin: !!args.options.skipRecycleBin,
            fromRecycleBin: !!args.options.fromRecycleBin,
            force: !!args.options.force
        });
    });
}, _SpoSiteRemoveCommand_initOptions = function _SpoSiteRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    }, {
        option: '--skipRecycleBin'
    }, {
        option: '--fromRecycleBin'
    }, {
        option: '-f, --force'
    });
}, _SpoSiteRemoveCommand_initValidators = function _SpoSiteRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.url);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        const uri = new URL(args.options.url);
        const rootUrl = `${uri.protocol}//${uri.hostname}`;
        if (rootUrl.toLowerCase() === urlUtil.removeTrailingSlashes(args.options.url.toLowerCase())) {
            return `The root site cannot be deleted.`;
        }
        if (args.options.fromRecycleBin && args.options.skipRecycleBin) {
            return 'Specify either fromRecycleBin or skipRecycleBin, but not both.';
        }
        return true;
    });
}, _SpoSiteRemoveCommand_initTypes = function _SpoSiteRemoveCommand_initTypes() {
    this.types.string.push('url');
    this.types.boolean.push('skipRecycleBin', 'fromRecycleBin', 'force');
};
export default new SpoSiteRemoveCommand();
//# sourceMappingURL=site-remove.js.map