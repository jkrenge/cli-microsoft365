var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemAttachmentListCommand_instances, _SpoListItemAttachmentListCommand_initTelemetry, _SpoListItemAttachmentListCommand_initOptions, _SpoListItemAttachmentListCommand_initValidators, _SpoListItemAttachmentListCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemAttachmentListCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_ATTACHMENT_LIST;
    }
    get description() {
        return 'Gets the attachments associated to a list item';
    }
    constructor() {
        super();
        _SpoListItemAttachmentListCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentListCommand_instances, "m", _SpoListItemAttachmentListCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentListCommand_instances, "m", _SpoListItemAttachmentListCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentListCommand_instances, "m", _SpoListItemAttachmentListCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentListCommand_instances, "m", _SpoListItemAttachmentListCommand_initOptionSets).call(this);
    }
    defaultProperties() {
        return ['FileName', 'ServerRelativeUrl'];
    }
    async commandAction(logger, args) {
        let requestUrl = `${args.options.webUrl}/_api/web`;
        if (args.options.listId) {
            requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
        }
        else if (args.options.listTitle) {
            requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
        }
        else if (args.options.listUrl) {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
            requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
        const requestOptions = {
            url: `${requestUrl}/items(${args.options.listItemId})?$select=AttachmentFiles&$expand=AttachmentFiles`,
            method: 'GET',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const attachmentFiles = await request.get(requestOptions);
            await logger.log(attachmentFiles.AttachmentFiles);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListItemAttachmentListCommand_instances = new WeakSet(), _SpoListItemAttachmentListCommand_initTelemetry = function _SpoListItemAttachmentListCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListItemAttachmentListCommand_initOptions = function _SpoListItemAttachmentListCommand_initOptions() {
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
}, _SpoListItemAttachmentListCommand_initValidators = function _SpoListItemAttachmentListCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        if (isNaN(parseInt(args.options.listItemId))) {
            return `${args.options.listItemId} is not a number`;
        }
        return true;
    });
}, _SpoListItemAttachmentListCommand_initOptionSets = function _SpoListItemAttachmentListCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemAttachmentListCommand();
//# sourceMappingURL=listitem-attachment-list.js.map