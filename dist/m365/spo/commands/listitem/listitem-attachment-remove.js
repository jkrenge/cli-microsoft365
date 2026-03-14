var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemAttachmentRemoveCommand_instances, _SpoListItemAttachmentRemoveCommand_initTelemetry, _SpoListItemAttachmentRemoveCommand_initOptions, _SpoListItemAttachmentRemoveCommand_initValidators, _SpoListItemAttachmentRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoListItemAttachmentRemoveCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_ATTACHMENT_REMOVE;
    }
    get description() {
        return 'Removes an attachment from a list item';
    }
    constructor() {
        super();
        _SpoListItemAttachmentRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentRemoveCommand_instances, "m", _SpoListItemAttachmentRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentRemoveCommand_instances, "m", _SpoListItemAttachmentRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentRemoveCommand_instances, "m", _SpoListItemAttachmentRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemAttachmentRemoveCommand_instances, "m", _SpoListItemAttachmentRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        const removeListItemAttachment = async () => {
            if (this.verbose) {
                const list = (args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl);
                await logger.logToStderr(`Removing attachment ${args.options.fileName} of item with id ${args.options.listItemId} from list ${list} in site at ${args.options.webUrl}...`);
            }
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
                headers: {
                    'X-HTTP-Method': 'DELETE',
                    'If-Match': '*',
                    'accept': 'application/json;odata=nometadata'
                },
                responseType: 'json'
            };
            try {
                await request.post(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeListItemAttachment();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the attachment ${args.options.fileName} of item with id ${args.options.listItemId} from the list ${args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl} in site ${args.options.webUrl}?` });
            if (result) {
                await removeListItemAttachment();
            }
        }
    }
}
_SpoListItemAttachmentRemoveCommand_instances = new WeakSet(), _SpoListItemAttachmentRemoveCommand_initTelemetry = function _SpoListItemAttachmentRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            force: (!(!args.options.force)).toString()
        });
    });
}, _SpoListItemAttachmentRemoveCommand_initOptions = function _SpoListItemAttachmentRemoveCommand_initOptions() {
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
    }, {
        option: '-f, --force'
    });
}, _SpoListItemAttachmentRemoveCommand_initValidators = function _SpoListItemAttachmentRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} in option listId is not a valid GUID`;
        }
        if (isNaN(args.options.listItemId)) {
            return `${args.options.listItemId} is not a number`;
        }
        return true;
    });
}, _SpoListItemAttachmentRemoveCommand_initOptionSets = function _SpoListItemAttachmentRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemAttachmentRemoveCommand();
//# sourceMappingURL=listitem-attachment-remove.js.map