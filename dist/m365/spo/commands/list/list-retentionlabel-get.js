var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListRetentionLabelGetCommand_instances, _SpoListRetentionLabelGetCommand_initTelemetry, _SpoListRetentionLabelGetCommand_initOptions, _SpoListRetentionLabelGetCommand_initValidators, _SpoListRetentionLabelGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListRetentionLabelGetCommand extends SpoCommand {
    get name() {
        return commands.LIST_RETENTIONLABEL_GET;
    }
    get description() {
        return 'Gets the default retention label set on the specified list or library.';
    }
    constructor() {
        super();
        _SpoListRetentionLabelGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelGetCommand_instances, "m", _SpoListRetentionLabelGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelGetCommand_instances, "m", _SpoListRetentionLabelGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelGetCommand_instances, "m", _SpoListRetentionLabelGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListRetentionLabelGetCommand_instances, "m", _SpoListRetentionLabelGetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        try {
            if (this.verbose) {
                await logger.logToStderr(`Getting label set on the list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
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
                    requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')?$expand=RootFolder&$select=RootFolder`;
                }
                else if (args.options.listTitle) {
                    if (this.debug) {
                        await logger.logToStderr(`Retrieving List from Title '${args.options.listTitle}'...`);
                    }
                    requestUrl += `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')?$expand=RootFolder&$select=RootFolder`;
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
                url: `${args.options.webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`,
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
            if (res['odata.null'] !== true) {
                await logger.log(res);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListRetentionLabelGetCommand_instances = new WeakSet(), _SpoListRetentionLabelGetCommand_initTelemetry = function _SpoListRetentionLabelGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: (!(!args.options.listId)).toString(),
            listTitle: (!(!args.options.listTitle)).toString(),
            listUrl: (!(!args.options.listUrl)).toString()
        });
    });
}, _SpoListRetentionLabelGetCommand_initOptions = function _SpoListRetentionLabelGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListRetentionLabelGetCommand_initValidators = function _SpoListRetentionLabelGetCommand_initValidators() {
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
}, _SpoListRetentionLabelGetCommand_initOptionSets = function _SpoListRetentionLabelGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListRetentionLabelGetCommand();
//# sourceMappingURL=list-retentionlabel-get.js.map