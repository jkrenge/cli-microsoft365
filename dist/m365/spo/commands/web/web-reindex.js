var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoWebReindexCommand_instances, _SpoWebReindexCommand_initOptions, _SpoWebReindexCommand_initValidators;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SpoPropertyBagBaseCommand } from '../propertybag/propertybag-base.js';
class SpoWebReindexCommand extends SpoCommand {
    get name() {
        return commands.WEB_REINDEX;
    }
    get description() {
        return 'Requests reindexing the specified subsite';
    }
    constructor() {
        super();
        _SpoWebReindexCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoWebReindexCommand_instances, "m", _SpoWebReindexCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoWebReindexCommand_instances, "m", _SpoWebReindexCommand_initValidators).call(this);
    }
    async commandAction(logger, args) {
        let webIdentityResp;
        try {
            const res = await spo.getRequestDigest(args.options.url);
            const requestDigest = res.FormDigestValue;
            if (this.debug) {
                await logger.logToStderr(`Retrieved request digest. Retrieving web identity...`);
            }
            const identityResp = await spo.getCurrentWebIdentity(args.options.url, requestDigest);
            webIdentityResp = identityResp;
            if (this.debug) {
                await logger.logToStderr(`Retrieved web identity.`);
            }
            if (this.verbose) {
                await logger.logToStderr(`Checking if the site is a no-script site...`);
            }
            const isNoScriptSite = await SpoPropertyBagBaseCommand.isNoScriptSite(args.options.url, requestDigest, webIdentityResp, logger, this.debug);
            if (isNoScriptSite) {
                if (this.verbose) {
                    await logger.logToStderr(`Site is a no-script site. Reindexing lists instead...`);
                }
                await this.reindexLists(args.options.url, requestDigest, logger, webIdentityResp);
            }
            if (this.verbose) {
                await logger.logToStderr(`Site is not a no-script site. Reindexing site...`);
            }
            const requestOptions = {
                url: `${args.options.url}/_api/web/allproperties`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            const webProperties = await request.get(requestOptions);
            let searchVersion = webProperties.vti_x005f_searchversion || 0;
            searchVersion++;
            await SpoPropertyBagBaseCommand.setProperty('vti_searchversion', searchVersion.toString(), args.options.url, requestDigest, webIdentityResp, logger, this.debug);
        }
        catch (err) {
            this.handleRejectedPromise(err);
        }
    }
    async reindexLists(webUrl, requestDigest, logger, webIdentityResp) {
        if (this.debug) {
            await logger.logToStderr(`Retrieving information about lists...`);
        }
        const requestOptions = {
            url: `${webUrl}/_api/web/lists?$select=NoCrawl,Title,RootFolder/Properties,RootFolder/ServerRelativeUrl&$expand=RootFolder/Properties`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const lists = await request.get(requestOptions);
        const promises = lists.value.map(l => this.reindexList(l, webUrl, requestDigest, webIdentityResp, logger));
        await Promise.all(promises);
    }
    async reindexList(list, webUrl, requestDigest, webIdentityResp, logger) {
        if (list.NoCrawl) {
            if (this.debug) {
                await logger.logToStderr(`List ${list.Title} is excluded from crawling`);
            }
            return;
        }
        const folderIdentityResp = await spo.getFolderIdentity(webIdentityResp.objectIdentity, webUrl, list.RootFolder.ServerRelativeUrl, requestDigest);
        let searchversion = list.RootFolder.Properties.vti_x005f_searchversion || 0;
        searchversion++;
        await SpoPropertyBagBaseCommand.setProperty('vti_searchversion', searchversion.toString(), webUrl, requestDigest, folderIdentityResp, logger, this.debug, list.RootFolder.ServerRelativeUrl);
        return;
    }
}
_SpoWebReindexCommand_instances = new WeakSet(), _SpoWebReindexCommand_initOptions = function _SpoWebReindexCommand_initOptions() {
    this.options.unshift({
        option: '-u, --url <url>'
    });
}, _SpoWebReindexCommand_initValidators = function _SpoWebReindexCommand_initValidators() {
    this.validators.push(async (args) => validation.isValidSharePointUrl(args.options.url));
};
export default new SpoWebReindexCommand();
//# sourceMappingURL=web-reindex.js.map