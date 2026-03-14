var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRecordLockCommand_instances, _SpoListItemRecordLockCommand_initTelemetry, _SpoListItemRecordLockCommand_initOptions, _SpoListItemRecordLockCommand_initValidators, _SpoListItemRecordLockCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemRecordLockCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_RECORD_LOCK;
    }
    get description() {
        return 'Locks the list item record';
    }
    constructor() {
        super();
        _SpoListItemRecordLockCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRecordLockCommand_instances, "m", _SpoListItemRecordLockCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordLockCommand_instances, "m", _SpoListItemRecordLockCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordLockCommand_instances, "m", _SpoListItemRecordLockCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordLockCommand_instances, "m", _SpoListItemRecordLockCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Locking the list item record ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
        }
        try {
            let listRestUrl = '';
            let listServerRelativeUrl = '';
            if (args.options.listUrl) {
                const listServerRelativeUrlFromPath = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
                listServerRelativeUrl = listServerRelativeUrlFromPath;
            }
            else {
                if (args.options.listId) {
                    listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
                }
                else {
                    listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
                }
                const requestOptions = {
                    url: `${args.options.webUrl}/_api/web/${listRestUrl}?$expand=RootFolder&$select=RootFolder`,
                    headers: {
                        'accept': 'application/json;odata=nometadata'
                    },
                    responseType: 'json'
                };
                const listInstance = await request.get(requestOptions);
                listServerRelativeUrl = listInstance.RootFolder.ServerRelativeUrl;
            }
            const listAbsoluteUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
            const requestUrl = `${args.options.webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.LockRecordItem()`;
            const requestOptions = {
                url: requestUrl,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: {
                    listUrl: listAbsoluteUrl,
                    itemId: args.options.listItemId
                },
                responseType: 'json'
            };
            await request.post(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListItemRecordLockCommand_instances = new WeakSet(), _SpoListItemRecordLockCommand_initTelemetry = function _SpoListItemRecordLockCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListItemRecordLockCommand_initOptions = function _SpoListItemRecordLockCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listItemId <listItemId>'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    });
}, _SpoListItemRecordLockCommand_initValidators = function _SpoListItemRecordLockCommand_initValidators() {
    this.validators.push(async (args) => {
        const id = parseInt(args.options.listItemId);
        if (isNaN(id)) {
            return `${args.options.listItemId} is not a valid list item ID`;
        }
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId &&
            !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoListItemRecordLockCommand_initOptionSets = function _SpoListItemRecordLockCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemRecordLockCommand();
//# sourceMappingURL=listitem-record-lock.js.map