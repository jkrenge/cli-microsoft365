var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemAttachmentAddCommand_instances, _SpoListItemAttachmentAddCommand_initTelemetry, _SpoListItemAttachmentAddCommand_initOptions, _SpoListItemAttachmentAddCommand_initValidators, _SpoListItemAttachmentAddCommand_initOptionSets;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import fs from 'fs';
class SpoListItemAttachmentAddCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_ATTACHMENT_ADD;
    }
    get description() {
        return 'Adds an attachment to a list item';
    }
    constructor() {
        super();
        _SpoListItemAttachmentAddCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentAddCommand_instances, "m", _SpoListItemAttachmentAddCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentAddCommand_instances, "m", _SpoListItemAttachmentAddCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentAddCommand_instances, "m", _SpoListItemAttachmentAddCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentAddCommand_instances, "m", _SpoListItemAttachmentAddCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Adding an attachment to list item with id ${args.options.listItemId} on list ${args.options.listId || args.options.listTitle || args.options.listUrl} on web ${args.options.webUrl}.`);
        }
        try {
            const fileName = this.getFileName(args.options.filePath, args.options.fileName);
            const fileBody = fs.readFileSync(args.options.filePath);
            const requestOptions = {
                url: `${args.options.webUrl}/_api/web/${this.getListUrl(args.options.webUrl, args.options.listId, args.options.listTitle, args.options.listUrl)}/items(${args.options.listItemId})/AttachmentFiles/add(FileName='${fileName}')`,
                headers: {
                    'accept': 'application/json;odata=nometadata'
                },
                data: fileBody,
                responseType: 'json'
            };
            const response = await request.post(requestOptions);
            await logger.log(response);
        }
        catch (err) {
            if (err.error &&
                err.error['odata.error'] &&
                err.error['odata.error'].message && err.error['odata.error'].message.value.indexOf('The document or folder name was not changed.') > -1) {
                this.handleError(err.error['odata.error'].message.value.split('\n')[0]);
            }
            else {
                this.handleRejectedODataJsonPromise(err);
            }
        }
    }
    getFileName(filePath, fileName) {
        if (!fileName) {
            return filePath.replace(/^.*[\\/]/, '');
        }
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
_SpoListItemAttachmentAddCommand_instances = new WeakSet(), _SpoListItemAttachmentAddCommand_initTelemetry = function _SpoListItemAttachmentAddCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            fileName: typeof args.options.fileName !== 'undefined'
        });
    });
}, _SpoListItemAttachmentAddCommand_initOptions = function _SpoListItemAttachmentAddCommand_initOptions() {
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
        option: '-n, --fileName [fileName]'
    });
}, _SpoListItemAttachmentAddCommand_initValidators = function _SpoListItemAttachmentAddCommand_initValidators() {
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
}, _SpoListItemAttachmentAddCommand_initOptionSets = function _SpoListItemAttachmentAddCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemAttachmentAddCommand();
//# sourceMappingURL=listitem-attachment-add.js.map