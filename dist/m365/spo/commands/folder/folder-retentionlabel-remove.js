var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderRetentionLabelRemoveCommand_instances, _SpoFolderRetentionLabelRemoveCommand_initTelemetry, _SpoFolderRetentionLabelRemoveCommand_initOptions, _SpoFolderRetentionLabelRemoveCommand_initValidators, _SpoFolderRetentionLabelRemoveCommand_initOptionSets, _SpoFolderRetentionLabelRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderRetentionLabelRemoveCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_RETENTIONLABEL_REMOVE;
    }
    get description() {
        return 'Clear the retention label from a folder';
    }
    constructor() {
        super();
        _SpoFolderRetentionLabelRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelRemoveCommand_instances, "m", _SpoFolderRetentionLabelRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelRemoveCommand_instances, "m", _SpoFolderRetentionLabelRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelRemoveCommand_instances, "m", _SpoFolderRetentionLabelRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelRemoveCommand_instances, "m", _SpoFolderRetentionLabelRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelRemoveCommand_instances, "m", _SpoFolderRetentionLabelRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (args.options.force) {
            await this.removeFolderRetentionLabel(logger, args);
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the retentionlabel from folder ${args.options.folderId || args.options.folderUrl} located in site ${args.options.webUrl}?` });
            if (result) {
                await this.removeFolderRetentionLabel(logger, args);
            }
        }
    }
    async removeFolderRetentionLabel(logger, args) {
        try {
            const folderProperties = await this.getFolderProperties(logger, args);
            if (folderProperties.ListItemAllFields) {
                const parsedUrl = new URL(args.options.webUrl);
                const tenantUrl = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
                const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, folderProperties.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl);
                await spo.removeRetentionLabelFromListItems(args.options.webUrl, listAbsoluteUrl, [parseInt(folderProperties.ListItemAllFields.Id)], logger, args.options.verbose);
            }
            else {
                const listAbsoluteUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, folderProperties.ServerRelativeUrl);
                await spo.removeDefaultRetentionLabelFromList(args.options.webUrl, listAbsoluteUrl, logger, args.options.verbose);
            }
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
    async getFolderProperties(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving list and item information for folder '${args.options.folderId || args.options.folderUrl}' in site at ${args.options.webUrl}...`);
        }
        let requestUrl = `${args.options.webUrl}/_api/web/`;
        if (args.options.folderId) {
            requestUrl += `GetFolderById('${args.options.folderId}')`;
        }
        else {
            const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
            requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
        }
        const requestOptions = {
            url: `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList/RootFolder&$select=ServerRelativeUrl,ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl,ListItemAllFields/Id`,
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            responseType: 'json'
        };
        return await request.get(requestOptions);
    }
}
_SpoFolderRetentionLabelRemoveCommand_instances = new WeakSet(), _SpoFolderRetentionLabelRemoveCommand_initTelemetry = function _SpoFolderRetentionLabelRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folderUrl: typeof args.options.folderUrl !== 'undefined',
            folderId: typeof args.options.folderId !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFolderRetentionLabelRemoveCommand_initOptions = function _SpoFolderRetentionLabelRemoveCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--folderUrl [folderUrl]'
    }, {
        option: '-i, --folderId [folderId]'
    }, {
        option: '-f, --force'
    });
}, _SpoFolderRetentionLabelRemoveCommand_initValidators = function _SpoFolderRetentionLabelRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.folderId &&
            !validation.isValidGuid(args.options.folderId)) {
            return `${args.options.folderId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoFolderRetentionLabelRemoveCommand_initOptionSets = function _SpoFolderRetentionLabelRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['folderUrl', 'folderId'] });
}, _SpoFolderRetentionLabelRemoveCommand_initTypes = function _SpoFolderRetentionLabelRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl', 'folderId');
    this.types.boolean.push('force');
};
export default new SpoFolderRetentionLabelRemoveCommand();
//# sourceMappingURL=folder-retentionlabel-remove.js.map