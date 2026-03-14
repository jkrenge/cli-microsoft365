var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderRetentionLabelEnsureCommand_instances, _SpoFolderRetentionLabelEnsureCommand_initTelemetry, _SpoFolderRetentionLabelEnsureCommand_initOptions, _SpoFolderRetentionLabelEnsureCommand_initValidators, _SpoFolderRetentionLabelEnsureCommand_initOptionSets, _SpoFolderRetentionLabelEnsureCommand_initTypes;
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderRetentionLabelEnsureCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_RETENTIONLABEL_ENSURE;
    }
    get description() {
        return 'Apply a retention label to a folder';
    }
    constructor() {
        super();
        _SpoFolderRetentionLabelEnsureCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelEnsureCommand_instances, "m", _SpoFolderRetentionLabelEnsureCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelEnsureCommand_instances, "m", _SpoFolderRetentionLabelEnsureCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelEnsureCommand_instances, "m", _SpoFolderRetentionLabelEnsureCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelEnsureCommand_instances, "m", _SpoFolderRetentionLabelEnsureCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderRetentionLabelEnsureCommand_instances, "m", _SpoFolderRetentionLabelEnsureCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        try {
            const folderProperties = await this.getFolderProperties(logger, args);
            if (folderProperties.ListItemAllFields) {
                const parsedUrl = new URL(args.options.webUrl);
                const tenantUrl = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
                const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, folderProperties.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl);
                await spo.applyRetentionLabelToListItems(args.options.webUrl, args.options.name, listAbsoluteUrl, [parseInt(folderProperties.ListItemAllFields.Id)], logger, args.options.verbose);
            }
            else {
                const listAbsoluteUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, folderProperties.ServerRelativeUrl);
                await spo.applyDefaultRetentionLabelToList(args.options.webUrl, args.options.name, listAbsoluteUrl, false, logger, args.options.verbose);
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
_SpoFolderRetentionLabelEnsureCommand_instances = new WeakSet(), _SpoFolderRetentionLabelEnsureCommand_initTelemetry = function _SpoFolderRetentionLabelEnsureCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folderUrl: typeof args.options.folderUrl !== 'undefined',
            folderId: typeof args.options.folderId !== 'undefined'
        });
    });
}, _SpoFolderRetentionLabelEnsureCommand_initOptions = function _SpoFolderRetentionLabelEnsureCommand_initOptions() {
    this.options.unshift({
        option: '-u, --webUrl <webUrl>'
    }, {
        option: '--name <name>'
    }, {
        option: '--folderUrl [folderUrl]'
    }, {
        option: '-i, --folderId [folderId]'
    });
}, _SpoFolderRetentionLabelEnsureCommand_initValidators = function _SpoFolderRetentionLabelEnsureCommand_initValidators() {
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
}, _SpoFolderRetentionLabelEnsureCommand_initOptionSets = function _SpoFolderRetentionLabelEnsureCommand_initOptionSets() {
    this.optionSets.push({ options: ['folderUrl', 'folderId'] });
}, _SpoFolderRetentionLabelEnsureCommand_initTypes = function _SpoFolderRetentionLabelEnsureCommand_initTypes() {
    this.types.string.push('webUrl', 'name', 'folderUrl', 'folderId');
};
export default new SpoFolderRetentionLabelEnsureCommand();
//# sourceMappingURL=folder-retentionlabel-ensure.js.map