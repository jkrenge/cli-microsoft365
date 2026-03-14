var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoListItemRetentionLabelRemoveCommand_instances, _SpoListItemRetentionLabelRemoveCommand_initTelemetry, _SpoListItemRetentionLabelRemoveCommand_initOptions, _SpoListItemRetentionLabelRemoveCommand_initValidators, _SpoListItemRetentionLabelRemoveCommand_initOptionSets;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
class SpoListItemRetentionLabelRemoveCommand extends SpoCommand {
    get name() {
        return commands.LISTITEM_RETENTIONLABEL_REMOVE;
    }
    get description() {
        return 'Clear the retention label from a list item';
    }
    constructor() {
        super();
        _SpoListItemRetentionLabelRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoListItemRetentionLabelRemoveCommand_instances, "m", _SpoListItemRetentionLabelRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoListItemRetentionLabelRemoveCommand_instances, "m", _SpoListItemRetentionLabelRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoListItemRetentionLabelRemoveCommand_instances, "m", _SpoListItemRetentionLabelRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoListItemRetentionLabelRemoveCommand_instances, "m", _SpoListItemRetentionLabelRemoveCommand_initOptionSets).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeListItemRetentionLabel(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the retentionlabel from list item ${args.options.listItemId} from list '${args.options.listId || args.options.listTitle || args.options.listUrl}' located in site ${args.options.webUrl}?` });
            if (result) {
                await this.removeListItemRetentionLabel(logger, args);
            }
        }
    }
    async removeListItemRetentionLabel(logger, args) {
        try {
            const listAbsoluteUrl = await this.getListAbsoluteUrl(args.options, logger);
            await spo.removeRetentionLabelFromListItems(args.options.webUrl, listAbsoluteUrl, [parseInt(args.options.listItemId)], logger, args.options.verbose);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getListAbsoluteUrl(options, logger) {
        const parsedUrl = new URL(options.webUrl);
        const tenantUrl = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
        if (options.listUrl) {
            const serverRelativePath = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
            return urlUtil.urlCombine(tenantUrl, serverRelativePath);
        }
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list absolute URL...`);
        }
        let requestUrl = `${options.webUrl}/_api/web`;
        if (options.listId) {
            requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
        }
        else if (options.listTitle) {
            requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
        }
        const requestOptions = {
            url: `${requestUrl}?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        const serverRelativePath = urlUtil.getServerRelativePath(options.webUrl, response.RootFolder.ServerRelativeUrl);
        const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, serverRelativePath);
        if (this.verbose) {
            await logger.logToStderr(`List absolute URL found: '${listAbsoluteUrl}'`);
        }
        return listAbsoluteUrl;
    }
}
_SpoListItemRetentionLabelRemoveCommand_instances = new WeakSet(), _SpoListItemRetentionLabelRemoveCommand_initTelemetry = function _SpoListItemRetentionLabelRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            listId: typeof args.options.listId !== 'undefined',
            listTitle: typeof args.options.listTitle !== 'undefined',
            listUrl: typeof args.options.listUrl !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoListItemRetentionLabelRemoveCommand_initOptions = function _SpoListItemRetentionLabelRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '-i, --listItemId <listItemId>'
    }, {
        option: '-l, --listId [listId]'
    }, {
        option: '-t, --listTitle [listTitle]'
    }, {
        option: '--listUrl [listUrl]'
    }, {
        option: '-f, --force'
    });
}, _SpoListItemRetentionLabelRemoveCommand_initValidators = function _SpoListItemRetentionLabelRemoveCommand_initValidators() {
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
}, _SpoListItemRetentionLabelRemoveCommand_initOptionSets = function _SpoListItemRetentionLabelRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
};
export default new SpoListItemRetentionLabelRemoveCommand();
//# sourceMappingURL=listitem-retentionlabel-remove.js.map