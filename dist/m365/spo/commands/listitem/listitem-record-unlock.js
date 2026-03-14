var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRecordUnlockCommand_instances, _SpoListItemRecordUnlockCommand_initTelemetry, _SpoListItemRecordUnlockCommand_initOptions, _SpoListItemRecordUnlockCommand_initValidators, _SpoListItemRecordUnlockCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemRecordUnlockCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_RECORD_UNLOCK;
    }
    get description() {
        return 'Unlocks the list item record';
    }
    constructor() {
        super();
        _SpoListItemRecordUnlockCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRecordUnlockCommand_instances, "m", _SpoListItemRecordUnlockCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordUnlockCommand_instances, "m", _SpoListItemRecordUnlockCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordUnlockCommand_instances, "m", _SpoListItemRecordUnlockCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRecordUnlockCommand_instances, "m", _SpoListItemRecordUnlockCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Unlocking the list item record ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
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
            const requestUrl = `${args.options.webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem()`;
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
_SpoListItemRecordUnlockCommand_instances = new WeakSet(), _SpoListItemRecordUnlockCommand_initTelemetry = function _SpoListItemRecordUnlockCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListItemRecordUnlockCommand_initOptions = function _SpoListItemRecordUnlockCommand_initOptions() {
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
}, _SpoListItemRecordUnlockCommand_initValidators = function _SpoListItemRecordUnlockCommand_initValidators() {
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
}, _SpoListItemRecordUnlockCommand_initOptionSets = function _SpoListItemRecordUnlockCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemRecordUnlockCommand();
//# sourceMappingURL=listitem-record-unlock.js.map