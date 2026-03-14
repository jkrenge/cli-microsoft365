var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListSiteScriptGetCommand_instances, _SpoListSiteScriptGetCommand_initTelemetry, _SpoListSiteScriptGetCommand_initOptions, _SpoListSiteScriptGetCommand_initValidators, _SpoListSiteScriptGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListSiteScriptGetCommand extends SpoCommand {
    get name() {
        return commands.LIST_SITESCRIPT_GET;
    }
    get description() {
        return 'Extracts a site script from a SharePoint list';
    }
    constructor() {
        super();
        _SpoListSiteScriptGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListSiteScriptGetCommand_instances, "m", _SpoListSiteScriptGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListSiteScriptGetCommand_instances, "m", _SpoListSiteScriptGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListSiteScriptGetCommand_instances, "m", _SpoListSiteScriptGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListSiteScriptGetCommand_instances, "m", _SpoListSiteScriptGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Extracting Site Script from list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
            }
            let listServerRelativeUrl = '';
            if (args.options.listUrl) {
                if (this.debug) {
                    await logger.logToStderr(`Retrieving List from URL '${args.options.listUrl}'...`);
                }
                listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            }
            else {
                let requestUrl = `${args.options.webUrl}/_api/web/`;
                if (args.options.listId) {
                    if (this.debug) {
                        await logger.logToStderr(`Retrieving List from Id '${args.options.listId}'...`);
                    }
                    requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')?$expand=RootFolder`;
                }
                else if (args.options.listTitle) {
                    if (this.debug) {
                        await logger.logToStderr(`Retrieving List from Title '${args.options.listTitle}'...`);
                    }
                    requestUrl += `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')?$expand=RootFolder`;
                }
                const requestOptions = {
                    url: requestUrl,
                    headers: {
                        'accept': 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const listInstance = await request.get(requestOptions);
                listServerRelativeUrl = listInstance.RootFolder.ServerRelativeUrl;
            }
            const listAbsoluteUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
            const reqOptions = {
                url: `${args.options.webUrl}/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`,
                headers: {
                    'accept': 'application/json;odata=nometadata',
                    'content-type': 'application/json;odata=nometadata'
                },
                responseType: 'json',
                data: {
                    listUrl: listAbsoluteUrl
                }
            };
            const res = await request.post(reqOptions);
            const siteScript = res.value;
            if (!siteScript) {
                throw `An error has occurred, the site script could not be extracted from list '${args.options.listId || args.options.listTitle}'`;
            }
            await logger.log(siteScript);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListSiteScriptGetCommand_instances = new WeakSet(), _SpoListSiteScriptGetCommand_initTelemetry = function _SpoListSiteScriptGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: (!(!args.options.listId)).toString(),
            listTitle: (!(!args.options.listTitle)).toString(),
            listUrl: (!(!args.options.listUrl)).toString()
        });
    });
}, _SpoListSiteScriptGetCommand_initOptions = function _SpoListSiteScriptGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListSiteScriptGetCommand_initValidators = function _SpoListSiteScriptGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId) {
            if (!validation.isValidGuid(args.options.listId)) {
                return `${args.options.listId} is not a valid GUID`;
            }
        }
        return true;
    });
}, _SpoListSiteScriptGetCommand_initOptionSets = function _SpoListSiteScriptGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListSiteScriptGetCommand();
//# sourceMappingURL=list-sitescript-get.js.map