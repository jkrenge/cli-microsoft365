var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderSharingLinkGetCommand_instances, _SpoFolderSharingLinkGetCommand_initTelemetry, _SpoFolderSharingLinkGetCommand_initOptions, _SpoFolderSharingLinkGetCommand_initOptionSets, _SpoFolderSharingLinkGetCommand_initValidators, _SpoFolderSharingLinkGetCommand_initTypes;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { drive as driveItem } from '../../../../utils/drive.js';
class SpoFolderSharingLinkGetCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_SHARINGLINK_GET;
    }
    get description() {
        return 'Gets details about a specific sharing link on a folder';
    }
    constructor() {
        super();
        _SpoFolderSharingLinkGetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkGetCommand_instances, "m", _SpoFolderSharingLinkGetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkGetCommand_instances, "m", _SpoFolderSharingLinkGetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkGetCommand_instances, "m", _SpoFolderSharingLinkGetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkGetCommand_instances, "m", _SpoFolderSharingLinkGetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkGetCommand_instances, "m", _SpoFolderSharingLinkGetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Retrieving sharing link on folder ${args.options.folderId || args.options.folderUrl}...`);
        }
        try {
            const relFolderUrl = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId);
            const absoluteFolderUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
            const folderUrl = new URL(absoluteFolderUrl);
            const siteId = await spo.getSiteIdByMSGraph(args.options.webUrl);
            const drive = await driveItem.getDriveByUrl(siteId, folderUrl);
            const itemId = await driveItem.getDriveItemId(drive, folderUrl);
            const requestUrl = `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${itemId}/permissions/${args.options.id}`;
            const requestOptions = {
                url: requestUrl,
                headers: {
                    'accept': 'application/json;odata.metadata=none'
                },
                responseType: 'json'
            };
            const permission = await request.get(requestOptions);
            await logger.log(permission);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFolderSharingLinkGetCommand_instances = new WeakSet(), _SpoFolderSharingLinkGetCommand_initTelemetry = function _SpoFolderSharingLinkGetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folderUrl: typeof args.options.folderUrl !== 'undefined',
            folderId: typeof args.options.folderId !== 'undefined'
        });
    });
}, _SpoFolderSharingLinkGetCommand_initOptions = function _SpoFolderSharingLinkGetCommand_initOptions() {
    this.options.unshift({ option: '-u, --webUrl <webUrl>' }, { option: '--folderUrl [folderUrl]' }, { option: '--folderId [folderId]' }, { option: '-i, --id <id>' });
}, _SpoFolderSharingLinkGetCommand_initOptionSets = function _SpoFolderSharingLinkGetCommand_initOptionSets() {
    this.optionSets.push({ options: ['folderUrl', 'folderId'] });
}, _SpoFolderSharingLinkGetCommand_initValidators = function _SpoFolderSharingLinkGetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.folderId && !validation.isValidGuid(args.options.folderId)) {
            return `${args.options.folderId} is not a valid GUID for option 'folderId'.`;
        }
        if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID for option 'id'.`;
        }
        return true;
    });
}, _SpoFolderSharingLinkGetCommand_initTypes = function _SpoFolderSharingLinkGetCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl', 'folderId', 'id');
};
export default new SpoFolderSharingLinkGetCommand();
//# sourceMappingURL=folder-sharinglink-get.js.map