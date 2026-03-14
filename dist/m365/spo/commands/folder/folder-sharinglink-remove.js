var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderSharingLinkRemoveCommand_instances, _SpoFolderSharingLinkRemoveCommand_initTelemetry, _SpoFolderSharingLinkRemoveCommand_initOptions, _SpoFolderSharingLinkRemoveCommand_initOptionSets, _SpoFolderSharingLinkRemoveCommand_initValidators, _SpoFolderSharingLinkRemoveCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import request from '../../../../request.js';
import commands from '../../commands.js';
class SpoFolderSharingLinkRemoveCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_SHARINGLINK_REMOVE;
    }
    get description() {
        return 'Removes a sharing link from a folder';
    }
    constructor() {
        super();
        _SpoFolderSharingLinkRemoveCommand_instances.add(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkRemoveCommand_instances, "m", _SpoFolderSharingLinkRemoveCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkRemoveCommand_instances, "m", _SpoFolderSharingLinkRemoveCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkRemoveCommand_instances, "m", _SpoFolderSharingLinkRemoveCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkRemoveCommand_instances, "m", _SpoFolderSharingLinkRemoveCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkRemoveCommand_instances, "m", _SpoFolderSharingLinkRemoveCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const removeSharingLink = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Removing sharing link of folder ${args.options.folderId || args.options.folderUrl} with id ${args.options.id}...`);
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
                        accept: 'application/json;odata.metadata=none'
                    },
                    responseType: 'json'
                };
                return request.delete(requestOptions);
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await removeSharingLink();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove sharing link ${args.options.id} of folder ${args.options.folderUrl || args.options.folderId}?` });
            if (result) {
                await removeSharingLink();
            }
        }
    }
}
_SpoFolderSharingLinkRemoveCommand_instances = new WeakSet(), _SpoFolderSharingLinkRemoveCommand_initTelemetry = function _SpoFolderSharingLinkRemoveCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            folderUrl: typeof args.options.folderUrl !== 'undefined',
            folderId: typeof args.options.folderId !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFolderSharingLinkRemoveCommand_initOptions = function _SpoFolderSharingLinkRemoveCommand_initOptions() {
    this.options.unshift({ option: '-u, --webUrl <webUrl>' }, { option: '--folderUrl [folderUrl]' }, { option: '--folderId [folderId]' }, { option: '-i, --id <id>' }, { option: '-f, --force' });
}, _SpoFolderSharingLinkRemoveCommand_initOptionSets = function _SpoFolderSharingLinkRemoveCommand_initOptionSets() {
    this.optionSets.push({ options: ['folderUrl', 'folderId'] });
}, _SpoFolderSharingLinkRemoveCommand_initValidators = function _SpoFolderSharingLinkRemoveCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.folderId && !validation.isValidGuid(args.options.folderId)) {
            return `${args.options.folderId} is not a valid GUID`;
        }
        return true;
    });
}, _SpoFolderSharingLinkRemoveCommand_initTypes = function _SpoFolderSharingLinkRemoveCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl', 'folderId', 'id');
    this.types.boolean.push('force');
};
export default new SpoFolderSharingLinkRemoveCommand();
//# sourceMappingURL=folder-sharinglink-remove.js.map