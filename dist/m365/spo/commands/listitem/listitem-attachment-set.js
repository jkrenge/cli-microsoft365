var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemAttachmentSetCommand_instances, _SpoListItemAttachmentSetCommand_initTelemetry, _SpoListItemAttachmentSetCommand_initOptions, _SpoListItemAttachmentSetCommand_initValidators, _SpoListItemAttachmentSetCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import fs from 'fs';
class SpoListItemAttachmentSetCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_ATTACHMENT_SET;
    }
    get description() {
        return 'Updates an attachment from a list item';
    }
    constructor() {
        super();
        _SpoListItemAttachmentSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentSetCommand_instances, "m", _SpoListItemAttachmentSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentSetCommand_instances, "m", _SpoListItemAttachmentSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentSetCommand_instances, "m", _SpoListItemAttachmentSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentSetCommand_instances, "m", _SpoListItemAttachmentSetCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Updating attachment ${args.options.fileName} at path ${args.options.filePath} for list item with id ${args.options.listItemId} on list ${args.options.listId || args.options.listTitle || args.options.listUrl} on web ${args.options.webUrl}.`);
        }
        try {
            const fileName = this.getFileName(args.options.filePath, args.options.fileName);
            const fileBody = fs.readFileSync(args.options.filePath);
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/${this.getListUrl(args.options.webUrl, args.options.listId, args.options.listTitle, args.options.listUrl)}/items(${args.options.listItemId})/AttachmentFiles('${fileName}')/$value`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: fileBody,
                responseType: 'json'
            };
            await request.put(requestOptions);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    getFileName(filePath, fileName) {
        const extension = filePath.split('.').pop();
        if (!fileName.endsWith(`.${extension}`)) {
            fileName += `.${extension}`;
        }
        return fileName;
    }
    getListUrl(webUrl, listId, listTitle, listUrl) {
        if (listId) {
            return `lists(guid'${formatting.encodeQueryParameter(listId)}')`;
        }
        else if (listTitle) {
            return `lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')`;
        }
        else {
            const listServerRelativeUrl = urlUtil.getServerRelativePath(webUrl, listUrl);
            return `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
        }
    }
}
_SpoListItemAttachmentSetCommand_instances = new WeakSet(), _SpoListItemAttachmentSetCommand_initTelemetry = function _SpoListItemAttachmentSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined'
        });
    });
}, _SpoListItemAttachmentSetCommand_initOptions = function _SpoListItemAttachmentSetCommand_initOptions() {
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
        option: '-p, --filePath <filePath>'
    }, {
        option: '-n, --fileName <fileName>'
    });
}, _SpoListItemAttachmentSetCommand_initValidators = function _SpoListItemAttachmentSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (isNaN(args.options.listItemId)) {
            return `${args.options.listItemId} in option listItemId is not a valid number.`;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID.`;
        }
        if (!fs.existsSync(args.options.filePath)) {
            return `File with path '${args.options.filePath}' was not found.`;
        }
        return true;
    });
}, _SpoListItemAttachmentSetCommand_initOptionSets = function _SpoListItemAttachmentSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemAttachmentSetCommand();
//# sourceMappingURL=listitem-attachment-set.js.map