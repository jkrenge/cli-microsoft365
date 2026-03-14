var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemAttachmentGetCommand_instances, _SpoListItemAttachmentGetCommand_initTelemetry, _SpoListItemAttachmentGetCommand_initOptions, _SpoListItemAttachmentGetCommand_initValidators, _SpoListItemAttachmentGetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemAttachmentGetCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_ATTACHMENT_GET;
    }
    get description() {
        return 'Gets an attachment from a list item';
    }
    constructor() {
        super();
        _SpoListItemAttachmentGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentGetCommand_instances, "m", _SpoListItemAttachmentGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentGetCommand_instances, "m", _SpoListItemAttachmentGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentGetCommand_instances, "m", _SpoListItemAttachmentGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentGetCommand_instances, "m", _SpoListItemAttachmentGetCommand_initOptionSets).call(this);
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
            url: `${requestUrl}/items(${args.options.listItemId})/AttachmentFiles('${args.options.fileName}')`,
            method: 'GET',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        try {
            const attachmentFile = await request.get(requestOptions);
            await logger.log(attachmentFile);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoListItemAttachmentGetCommand_instances = new WeakSet(), _SpoListItemAttachmentGetCommand_initTelemetry = function _SpoListItemAttachmentGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListItemAttachmentGetCommand_initOptions = function _SpoListItemAttachmentGetCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--listId [listId]'
    }, {
        option: '--listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '--listItemId <listItemId>'
    }, {
        option: '-n, --fileName <fileName>'
    });
}, _SpoListItemAttachmentGetCommand_initValidators = function _SpoListItemAttachmentGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID.`;
        }
        if (isNaN(args.options.listItemId)) {
            return `${args.options.listItemId} is not a number.`;
        }
        return true;
    });
}, _SpoListItemAttachmentGetCommand_initOptionSets = function _SpoListItemAttachmentGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemAttachmentGetCommand();
//# sourceMappingURL=listitem-attachment-get.js.map