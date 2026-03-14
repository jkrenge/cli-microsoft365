var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderSharingLinkSetCommand_instances, _SpoFolderSharingLinkSetCommand_initTelemetry, _SpoFolderSharingLinkSetCommand_initOptions, _SpoFolderSharingLinkSetCommand_initValidators, _SpoFolderSharingLinkSetCommand_initOptionSets, _SpoFolderSharingLinkSetCommand_initTypes;
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
class SpoFolderSharingLinkSetCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_SHARINGLINK_SET;
    }
    get description() {
        return 'Updates a specific sharing link to a folder';
    }
    constructor() {
        super();
        _SpoFolderSharingLinkSetCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkSetCommand_instances, "m", _SpoFolderSharingLinkSetCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkSetCommand_instances, "m", _SpoFolderSharingLinkSetCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkSetCommand_instances, "m", _SpoFolderSharingLinkSetCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkSetCommand_instances, "m", _SpoFolderSharingLinkSetCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkSetCommand_instances, "m", _SpoFolderSharingLinkSetCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        if (this.verbose) {
            await logger.logToStderr(`Updating sharing link to a folder ${args.options.folderId || args.options.folderUrl}...`);
        }
        try {
            const relFolderUrl = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId, logger, args.options.verbose);
            const absoluteFolderUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
            const folderUrl = new URL(absoluteFolderUrl);
            const siteId = await spo.getSiteIdByMSGraph(args.options.webUrl);
            const driveDetails = await drive.getDriveByUrl(siteId, folderUrl, logger, args.options.verbose);
            const itemId = await drive.getDriveItemId(driveDetails, folderUrl, logger, args.options.verbose);
            const requestOptions = {
                url: `https://graph.microsoft.com/v1.0/drives/${driveDetails.id}/items/${itemId}/permissions/${args.options.id}`,
                headers: {
                    accept: 'application/json;odata.metadata=none',
                    'content-type': 'application/json'
                },
                responseType: 'json',
                data: {
                    expirationDateTime: args.options.expirationDateTime
                }
            };
            const sharingLink = await request.patch(requestOptions);
            await logger.log(sharingLink);
        }
        catch (err) {
            this.handleRejectedODataJsonPromise(err);
        }
    }
}
_SpoFolderSharingLinkSetCommand_instances = new WeakSet(), _SpoFolderSharingLinkSetCommand_initTelemetry = function _SpoFolderSharingLinkSetCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folderUrl: typeof args.options.folderUrl !== 'undefined',
            folderId: typeof args.options.folderId !== 'undefined',
            expirationDateTime: typeof args.options.expirationDateTime !== 'undefined'
        });
    });
}, _SpoFolderSharingLinkSetCommand_initOptions = function _SpoFolderSharingLinkSetCommand_initOptions() {
    this.options.unshift({ option: '-u, --webUrl <webUrl>' }, { option: '--folderUrl [folderUrl]' }, { option: '--folderId [folderId]' }, { option: '-i, --id <id>' }, { option: '--expirationDateTime [expirationDateTime]' });
}, _SpoFolderSharingLinkSetCommand_initValidators = function _SpoFolderSharingLinkSetCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.folderId && !validation.isValidGuid(args.options.folderId)) {
            return `${args.options.folderId} is not a valid GUID`;
        }
        if (args.options.expirationDateTime && !validation.isValidISODateTime(args.options.expirationDateTime)) {
            return `${args.options.expirationDateTime} is not a valid ISO date string.`;
        }
        return true;
    });
}, _SpoFolderSharingLinkSetCommand_initOptionSets = function _SpoFolderSharingLinkSetCommand_initOptionSets() {
    this.optionSets.push({ options: ['folderId', 'folderUrl'] });
}, _SpoFolderSharingLinkSetCommand_initTypes = function _SpoFolderSharingLinkSetCommand_initTypes() {
    this.types.string.push('webUrl', 'folderId', 'folderUrl', 'id', 'expirationDateTime');
};
export default new SpoFolderSharingLinkSetCommand();
//# sourceMappingURL=folder-sharinglink-set.js.map