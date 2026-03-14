var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SpoFolderSharingLinkClearCommand_instances, _SpoFolderSharingLinkClearCommand_initTelemetry, _SpoFolderSharingLinkClearCommand_initOptions, _SpoFolderSharingLinkClearCommand_initOptionSets, _SpoFolderSharingLinkClearCommand_initValidators, _SpoFolderSharingLinkClearCommand_initTypes;
import { cli } from '../../../../cli/cli.js';
import { spo } from '../../../../utils/spo.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import request from '../../../../request.js';
import commands from '../../commands.js';
class SpoFolderSharingLinkClearCommand extends SpoCommand {
    get name() {
        return commands.FOLDER_SHARINGLINK_CLEAR;
    }
    get description() {
        return 'Removes all sharing links of a folder';
    }
    constructor() {
        super();
        _SpoFolderSharingLinkClearCommand_instances.add(this);
        this.allowedScopes = ['anonymous', 'users', 'organization'];
        __classPrivateFieldGet(this, _SpoFolderSharingLinkClearCommand_instances, "m", _SpoFolderSharingLinkClearCommand_initTelemetry).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkClearCommand_instances, "m", _SpoFolderSharingLinkClearCommand_initOptions).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkClearCommand_instances, "m", _SpoFolderSharingLinkClearCommand_initOptionSets).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkClearCommand_instances, "m", _SpoFolderSharingLinkClearCommand_initValidators).call(this);
        __classPrivateFieldGet(this, _SpoFolderSharingLinkClearCommand_instances, "m", _SpoFolderSharingLinkClearCommand_initTypes).call(this);
    }
    async commandAction(logger, args) {
        const clearSharingLinks = async () => {
            if (this.verbose) {
                await logger.logToStderr(`Clearing sharing links from folder ${args.options.folderId || args.options.folderUrl} for ${args.options.scope ? `${args.options.scope} scope` : 'all scopes'}`);
            }
            try {
                const relFolderUrl = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId, logger, args.options.verbose);
                const absoluteFolderUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
                const folderUrl = new URL(absoluteFolderUrl);
                const siteId = await spo.getSiteIdByMSGraph(args.options.webUrl);
                const driveDetails = await drive.getDriveByUrl(siteId, folderUrl, logger, args.options.verbose);
                const itemId = await drive.getDriveItemId(driveDetails, folderUrl, logger, args.options.verbose);
                let requestUrl = `https://graph.microsoft.com/v1.0/drives/${driveDetails.id}/items/${itemId}/permissions?$filter=Link ne null`;
                if (args.options.scope) {
                    requestUrl += ` and Link/Scope eq '${args.options.scope}'`;
                }
                const sharingLinks = await odata.getAllItems(requestUrl);
                for (const sharingLink of sharingLinks) {
                    const requestOptions = {
                        url: `https://graph.microsoft.com/v1.0/drives/${driveDetails.id}/items/${itemId}/permissions/${sharingLink.id}`,
                        headers: {
                            accept: 'application/json;odata.metadata=none'
                        }
                    };
                    await request.delete(requestOptions);
                }
            }
            catch (err) {
                this.handleRejectedODataJsonPromise(err);
            }
        };
        if (args.options.force) {
            await clearSharingLinks();
        }
        else {
            const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear sharing links from folder ${args.options.folderUrl || args.options.folderId}? for ${args.options.scope ? `${args.options.scope} scope` : 'all scopes'}` });
            if (result) {
                await clearSharingLinks();
            }
        }
    }
}
_SpoFolderSharingLinkClearCommand_instances = new WeakSet(), _SpoFolderSharingLinkClearCommand_initTelemetry = function _SpoFolderSharingLinkClearCommand_initTelemetry() {
    this.telemetry.push((args) => {
        Object.assign(this.telemetryProperties, {
            webUrl: typeof args.options.webUrl !== 'undefined',
            folderUrl: typeof args.options.folderUrl !== 'undefined',
            folderId: typeof args.options.folderId !== 'undefined',
            scope: typeof args.options.scope !== 'undefined',
            force: !!args.options.force
        });
    });
}, _SpoFolderSharingLinkClearCommand_initOptions = function _SpoFolderSharingLinkClearCommand_initOptions() {
    this.options.unshift({ option: '-u, --webUrl <webUrl>' }, { option: '--folderUrl [folderUrl]' }, { option: '--folderId [folderId]' }, {
        option: '--scope [scope]',
        autocomplete: this.allowedScopes
    }, { option: '-f, --force' });
}, _SpoFolderSharingLinkClearCommand_initOptionSets = function _SpoFolderSharingLinkClearCommand_initOptionSets() {
    this.optionSets.push({ options: ['folderUrl', 'folderId'] });
}, _SpoFolderSharingLinkClearCommand_initValidators = function _SpoFolderSharingLinkClearCommand_initValidators() {
    this.validators.push(async (args) => {
        const isValidSharePointUrl = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
        }
        if (args.options.folderId && !validation.isValidGuid(args.options.folderId)) {
            return `${args.options.folderId} is not a valid GUID`;
        }
        if (args.options.scope && !this.allowedScopes.some(scope => scope === args.options.scope)) {
            return `'${args.options.scope}' is not a valid scope. Allowed values are: ${this.allowedScopes.join(',')}`;
        }
        return true;
    });
}, _SpoFolderSharingLinkClearCommand_initTypes = function _SpoFolderSharingLinkClearCommand_initTypes() {
    this.types.string.push('webUrl', 'folderUrl', 'folderId', 'scope');
    this.types.boolean.push('force');
};
export default new SpoFolderSharingLinkClearCommand();
//# sourceMappingURL=folder-sharinglink-clear.js.map