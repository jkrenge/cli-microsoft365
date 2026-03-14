var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFileRetentionLabelRemoveCommand_instances, _SpoFileRetentionLabelRemoveCommand_initTelemetry, _SpoFileRetentionLabelRemoveCommand_initOptions, _SpoFileRetentionLabelRemoveCommand_initValidators, _SpoFileRetentionLabelRemoveCommand_initOptionSets, _SpoFileRetentionLabelRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFileRetentionLabelRemoveCommand extends SpoCommand {
    get name() {
        return commands.FILE_RETENTIONLABEL_REMOVE;
    }
    get description() {
        return 'Clear the retention label from a file';
    }
    constructor() {
        super();
        _SpoFileRetentionLabelRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelRemoveCommand_instances, "m", _SpoFileRetentionLabelRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelRemoveCommand_instances, "m", _SpoFileRetentionLabelRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelRemoveCommand_instances, "m", _SpoFileRetentionLabelRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelRemoveCommand_instances, "m", _SpoFileRetentionLabelRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFileRetentionLabelRemoveCommand_instances, "m", _SpoFileRetentionLabelRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeFileRetentionLabel(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the retentionlabel from file ${args.options.fileId || args.options.fileUrl} located in site ${args.options.webUrl}?` });
            if (result) {
                await this.removeFileRetentionLabel(logger, args);
            }
        }
    }
    async removeFileRetentionLabel(logger, args) {
        try {
            const fileProperties = await this.getFileProperties(logger, args);
            const parsedUrl = new URL(args.options.webUrl);
            const tenantUrl = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
            const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, fileProperties.listServerRelativeUrl);
            await spo.removeRetentionLabelFromListItems(args.options.webUrl, listAbsoluteUrl, [parseInt(fileProperties.listItemId)], logger, args.options.verbose);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getFileProperties(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list and item information for file '${args.options.fileId || args.options.fileUrl}' in site at ${args.options.webUrl}...`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.fileId) {
            requestUrl += `GetFileById('${args.options.fileId}')`;
        }
        else {
            const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
            requestUrl += `GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
        }
        const requestOptions = {
            url: `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList/RootFolder&$select=ServerRelativeUrl,ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl,ListItemAllFields/Id`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        const response = await request.get(requestOptions);
        return { listItemId: response.ListItemAllFields.Id, listServerRelativeUrl: response.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl };
    }
}
_SpoFileRetentionLabelRemoveCommand_instances = new WeakSet(), _SpoFileRetentionLabelRemoveCommand_initTelemetry = function _SpoFileRetentionLabelRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            fileUrl: typeof args.options.fileUrl !== 'undefined',
            fileId: typeof args.options.fileId !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFileRetentionLabelRemoveCommand_initOptions = function _SpoFileRetentionLabelRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--fileUrl [fileUrl]'
    }, {
        option: '-i, --fileId [fileId]'
    }, {
        option: '-f, --force'
    });
}, _SpoFileRetentionLabelRemoveCommand_initValidators = function _SpoFileRetentionLabelRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.fileId &&
            !validation.isValidGuid(args.options.fileId)) {
            return `${args.options.fileId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoFileRetentionLabelRemoveCommand_initOptionSets = function _SpoFileRetentionLabelRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
}, _SpoFileRetentionLabelRemoveCommand_initTypes = function _SpoFileRetentionLabelRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'fileUrl', 'fileId');
    this.types.boolean.push('force');
};
export default new SpoFileRetentionLabelRemoveCommand();
//# sourceMappingURL=file-retentionlabel-remove.js.map